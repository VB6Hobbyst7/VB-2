VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmManual 
   Caption         =   "검사(수작업접수)"
   ClientHeight    =   7800
   ClientLeft      =   240
   ClientTop       =   840
   ClientWidth     =   11850
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
   ScaleHeight     =   7800
   ScaleWidth      =   11850
   WindowState     =   2  '최대화
   Begin VB.CheckBox chkEx 
      Caption         =   "외부검사Check"
      Height          =   285
      Left            =   4770
      TabIndex        =   48
      Top             =   225
      Width           =   1545
   End
   Begin Threed.SSPanel panel_iD 
      Height          =   5550
      Left            =   225
      TabIndex        =   23
      Top             =   1935
      Width           =   4290
      _Version        =   65536
      _ExtentX        =   7567
      _ExtentY        =   9790
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
      Alignment       =   0
      Begin Threed.SSCommand cmdidClear 
         Height          =   375
         Left            =   2745
         TabIndex        =   43
         Top             =   4905
         Width           =   1410
         _Version        =   65536
         _ExtentX        =   2487
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Clear"
         BevelWidth      =   1
         Outline         =   0   'False
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   2205
         TabIndex        =   14
         Top             =   4410
         Width           =   1230
      End
      Begin VB.TextBox txtUserid 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1305
         TabIndex        =   13
         Top             =   4410
         Width           =   915
      End
      Begin MSComCtl2.DTPicker dtJeobsuT 
         Height          =   285
         Left            =   1305
         TabIndex        =   12
         Top             =   4095
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "hh:mm"
         Format          =   24772611
         UpDown          =   -1  'True
         CurrentDate     =   36365
      End
      Begin Threed.SSCommand cmdSample 
         Height          =   285
         Left            =   2160
         TabIndex        =   42
         Top             =   3780
         Width           =   600
         _Version        =   65536
         _ExtentX        =   1058
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Call"
      End
      Begin VB.TextBox txtSampleCode 
         Height          =   285
         Left            =   1305
         TabIndex        =   11
         Top             =   3780
         Width           =   825
      End
      Begin VB.TextBox txtIndate 
         Height          =   315
         Left            =   1305
         TabIndex        =   10
         Top             =   3420
         Width           =   1455
      End
      Begin VB.TextBox txtOrderno 
         Height          =   285
         Left            =   1305
         TabIndex        =   9
         Top             =   3105
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtOrderDt 
         Height          =   285
         Left            =   1305
         TabIndex        =   8
         Top             =   2790
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24772611
         CurrentDate     =   36365
      End
      Begin VB.TextBox txtRoom 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3015
         TabIndex        =   41
         Top             =   2430
         Width           =   870
      End
      Begin VB.TextBox txtioGubun 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3015
         TabIndex        =   40
         Top             =   2115
         Width           =   870
      End
      Begin VB.ComboBox cmbioGubun 
         Height          =   300
         ItemData        =   "frmManual.frx":0000
         Left            =   1305
         List            =   "frmManual.frx":000A
         Style           =   2  '드롭다운 목록
         TabIndex        =   7
         Top             =   2115
         Width           =   1680
      End
      Begin VB.TextBox txtDr 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3015
         TabIndex        =   39
         Top             =   1800
         Width           =   870
      End
      Begin VB.TextBox txtDept 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3015
         TabIndex        =   38
         Top             =   1485
         Width           =   870
      End
      Begin VB.ComboBox cmbDr 
         Height          =   300
         Left            =   1305
         Style           =   2  '드롭다운 목록
         TabIndex        =   6
         Top             =   1800
         Width           =   1680
      End
      Begin VB.ComboBox cmbDept 
         Height          =   300
         Left            =   1305
         Style           =   2  '드롭다운 목록
         TabIndex        =   5
         Top             =   1485
         Width           =   1680
      End
      Begin VB.TextBox txtAgemm 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   1170
         Width           =   510
      End
      Begin VB.TextBox txtAgeyy 
         Height          =   285
         Left            =   1305
         TabIndex        =   3
         Top             =   1170
         Width           =   465
      End
      Begin VB.TextBox txtSex 
         Height          =   285
         Left            =   1305
         TabIndex        =   2
         Top             =   855
         Width           =   465
      End
      Begin VB.TextBox txtSname 
         Height          =   285
         Left            =   1305
         TabIndex        =   1
         Top             =   540
         Width           =   1320
      End
      Begin VB.TextBox txtPtno 
         Height          =   285
         Left            =   1305
         TabIndex        =   0
         Top             =   225
         Width           =   1320
      End
      Begin VB.Label Label4 
         Caption         =   "입력후 Enter"
         Height          =   195
         Left            =   2700
         TabIndex        =   53
         Top             =   270
         Width           =   1140
      End
      Begin VB.Label labTitle 
         Caption         =   "접수자"
         Height          =   195
         Index           =   13
         Left            =   225
         TabIndex        =   37
         Top             =   4455
         Width           =   960
      End
      Begin VB.Label labTitle 
         Caption         =   "접수시간"
         Height          =   285
         Index           =   12
         Left            =   225
         TabIndex        =   36
         Top             =   4140
         Width           =   915
      End
      Begin VB.Label labTitle 
         Caption         =   "검체"
         Height          =   285
         Index           =   11
         Left            =   225
         TabIndex        =   35
         Top             =   3825
         Width           =   960
      End
      Begin VB.Label labTitle 
         Caption         =   "내원일자"
         Height          =   285
         Index           =   10
         Left            =   225
         TabIndex        =   34
         Top             =   3480
         Width           =   960
      End
      Begin VB.Label labTitle 
         Caption         =   "Order No."
         Height          =   285
         Index           =   9
         Left            =   225
         TabIndex        =   33
         Top             =   3150
         Width           =   960
      End
      Begin VB.Label labTitle 
         Caption         =   "Order일자"
         Height          =   285
         Index           =   8
         Left            =   225
         TabIndex        =   32
         Top             =   2820
         Width           =   960
      End
      Begin VB.Label labTitle 
         Caption         =   "병실"
         Height          =   240
         Index           =   7
         Left            =   2565
         TabIndex        =   31
         Top             =   2475
         Width           =   420
      End
      Begin VB.Label labTitle 
         Caption         =   "내원구분"
         Height          =   285
         Index           =   6
         Left            =   225
         TabIndex        =   30
         Top             =   2190
         Width           =   960
      End
      Begin VB.Label labTitle 
         Caption         =   "진료의사"
         Height          =   285
         Index           =   5
         Left            =   225
         TabIndex        =   29
         Top             =   1860
         Width           =   960
      End
      Begin VB.Label labTitle 
         Caption         =   "진료과"
         Height          =   285
         Index           =   4
         Left            =   225
         TabIndex        =   28
         Top             =   1515
         Width           =   960
      End
      Begin VB.Label labTitle 
         Caption         =   "나이"
         Height          =   285
         Index           =   3
         Left            =   225
         TabIndex        =   27
         Top             =   1230
         Width           =   960
      End
      Begin VB.Label labTitle 
         Caption         =   "성별"
         Height          =   285
         Index           =   2
         Left            =   225
         TabIndex        =   26
         Top             =   900
         Width           =   960
      End
      Begin VB.Label labTitle 
         Caption         =   "수진자명"
         Height          =   285
         Index           =   1
         Left            =   225
         TabIndex        =   25
         Top             =   600
         Width           =   960
      End
      Begin VB.Label labTitle 
         Caption         =   "병록번호"
         Height          =   285
         Index           =   0
         Left            =   225
         TabIndex        =   24
         Top             =   270
         Width           =   960
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1230
      Left            =   225
      TabIndex        =   15
      Top             =   675
      Width           =   4290
      _Version        =   65536
      _ExtentX        =   7567
      _ExtentY        =   2170
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      Begin Threed.SSCommand cmdGetno2 
         Height          =   285
         Left            =   2880
         TabIndex        =   51
         Top             =   810
         Width           =   240
         _Version        =   65536
         _ExtentX        =   423
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "g"
         Enabled         =   0   'False
      End
      Begin VB.TextBox txtSLipno2 
         Height          =   285
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   810
         Width           =   1095
      End
      Begin VB.TextBox txtSLipno1 
         Height          =   285
         Left            =   1305
         TabIndex        =   21
         Top             =   810
         Width           =   465
      End
      Begin MSComCtl2.DTPicker dtJeobsuDt 
         Height          =   330
         Left            =   1305
         TabIndex        =   18
         Top             =   450
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24772611
         CurrentDate     =   36365
      End
      Begin VB.ComboBox cmbSLip 
         Height          =   300
         Left            =   1305
         Style           =   2  '드롭다운 목록
         TabIndex        =   16
         Top             =   135
         Width           =   2805
      End
      Begin VB.Label Label3 
         Caption         =   "SLipno"
         Height          =   285
         Left            =   180
         TabIndex        =   20
         Top             =   855
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "접수일자"
         Height          =   240
         Left            =   180
         TabIndex        =   19
         Top             =   540
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "검사종류"
         Height          =   285
         Left            =   180
         TabIndex        =   17
         Top             =   180
         Width           =   1005
      End
   End
   Begin FPSpreadADO.fpSpread ssSlipList1 
      Height          =   3525
      Left            =   6525
      TabIndex        =   44
      Top             =   180
      Width           =   4200
      _Version        =   196608
      _ExtentX        =   7408
      _ExtentY        =   6218
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
      MaxCols         =   9
      MaxRows         =   200
      ScrollBars      =   2
      SpreadDesigner  =   "frmManual.frx":001E
      Appearance      =   1
      ScrollBarTrack  =   1
   End
   Begin FPSpreadADO.fpSpread ssSelectitem 
      Height          =   3720
      Left            =   4680
      TabIndex        =   45
      Top             =   3780
      Width           =   6045
      _Version        =   196608
      _ExtentX        =   10663
      _ExtentY        =   6562
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
      GridColor       =   8421440
      MaxCols         =   9
      MaxRows         =   50
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   12632256
      ShadowDark      =   8421504
      ShadowText      =   0
      SpreadDesigner  =   "frmManual.frx":1A47
      UserResize      =   0
      VisibleCols     =   7
      VisibleRows     =   50
      Appearance      =   1
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   420
      Left            =   225
      TabIndex        =   52
      Top             =   225
      Width           =   3300
      _Version        =   65536
      _ExtentX        =   5821
      _ExtentY        =   741
      _StockProps     =   15
      Caption         =   "수작업접수"
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
   Begin MSForms.CommandButton cmdEnrolOk 
      Height          =   510
      Left            =   4725
      TabIndex        =   50
      Top             =   3195
      Width           =   1635
      Caption         =   "접수확인"
      PicturePosition =   327683
      Size            =   "2884;900"
      Picture         =   "frmManual.frx":20F5
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdMove 
      Height          =   510
      Left            =   4770
      TabIndex        =   49
      Top             =   1980
      Visible         =   0   'False
      Width           =   1635
      Caption         =   "Expand"
      PicturePosition =   327683
      Size            =   "2884;900"
      Picture         =   "frmManual.frx":38B7
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdClear 
      Height          =   510
      Left            =   4770
      TabIndex        =   47
      Top             =   1260
      Width           =   1635
      Caption         =   "화면정리"
      PicturePosition =   327683
      Size            =   "2884;900"
      Picture         =   "frmManual.frx":4191
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdCodeSelect 
      Height          =   510
      Left            =   4770
      TabIndex        =   46
      Top             =   720
      Width           =   1635
      Caption         =   "코드조회"
      PicturePosition =   327683
      Size            =   "2884;900"
      Picture         =   "frmManual.frx":5923
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
Attribute VB_Name = "frmManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Function isPatientMaster(ByVal sFindPtno As String) As Integer
    Dim adoPt       As ADODB.Recordset
    
    strSql = " SELECT Ptno FROM TW_MIS_PMPA.TWBAS_PATIENT WHERE Ptno = '" & sFindPtno & "'"
    If False = adoSetOpen(strSql, adoPt) Then
        isPatientMaster = False
        Exit Function
    Else
        isPatientMaster = True
        Call adoSetClose(adoPt)
    End If
    
End Function



Private Sub chkEx_Click()
    
    If chkEx.Value = "1" Then
        chkEx.Tag = "W"
    Else
        chkEx.Tag = ""
    End If
    
End Sub

Private Sub cmbDept_Click()
    Dim adoDr       As ADODB.Recordset
    
    
    cmbDr.Clear
    If cmbDept.ListIndex = -1 Then Exit Sub
    txtDept.Text = Left(cmbDept, 4)
    
    strSql = ""
    strSql = strSql & " SELECT DrCode, Drname "
    strSql = strSql & " FROM   TW_MIS_PMPA.TWBAS_DOCTOR "
    strSql = strSql & " WHERE  DRDEPT1 = '" & txtDept.Text & "'"
    If adoSetOpen(strSql, adoDr) Then
        Do Until adoDr.EOF
            cmbDr.AddItem Trim(adoDr.Fields("DrCode").Value & "") & ". " & _
                          Trim(adoDr.Fields("Drname").Value & "")
            adoDr.MoveNext
        Loop
        Call adoSetClose(adoDr)
    End If
    txtDr.Text = ""
    
End Sub

Private Sub cmbDr_Click()
    
    If cmbDr.ListIndex = -1 Then Exit Sub
    
    txtDr.Text = Left(cmbDr.Text, 6)
    
End Sub

Private Sub cmbioGubun_Click()
    
    If cmbioGubun.ListIndex = -1 Then Exit Sub
    
    txtioGubun.Text = Left(cmbioGubun.Text, 1)
    
End Sub

Private Sub cmbSLip_Click()
    
    If cmbSLip.ListIndex = -1 Then Exit Sub
    
    txtSLipno1.Text = Left(cmbSLip.List(cmbSLip.ListIndex), 2)
    
    
End Sub

Private Sub cmdClear_Click()
    txtSLipno2.Text = ""
    Call cmdidClear_Click
    
    ssSlipList1.Row = 1
    ssSlipList1.Row2 = ssSlipList1.DataRowCnt
    ssSlipList1.Col = 1
    ssSlipList1.Col2 = ssSlipList1.DataColCnt
    ssSlipList1.BlockMode = True
    ssSlipList1.Action = ActionClear
    ssSlipList1.BlockMode = False
    
    ssSelectitem.Row = 1
    ssSelectitem.Row2 = ssSelectitem.DataRowCnt
    ssSelectitem.Col = 1
    ssSelectitem.Col2 = ssSelectitem.DataColCnt
    ssSelectitem.BlockMode = True
    ssSelectitem.Action = ActionClear
    ssSelectitem.BlockMode = False
    
    
End Sub

Private Sub cmdCodeSelect_Click()
    
    ssSlipList1.MaxRows = 200
    Call Spread_Set_Clear(ssSlipList1)
    
    GoSub SELECT_RoutinCode
    GoSub SELECT_ItemCode
    ssSlipList1.MaxRows = ssSlipList1.DataRowCnt
    
    Exit Sub
    


SELECT_RoutinCode:
    strSql = ""
    strSql = strSql & " SELECT  RoutinCd, RoutinNm "
    strSql = strSql & " FROM    TW_MIS_EXAM.TWEXAM_Routine "
    strSql = strSql & " WHERE   Codeky LIKE '" & Trim(txtSLipno1.Text) & "%'"
    strSql = strSql & " GROUP   BY RoutinCd, RoutinNm "
    strSql = strSql & " ORDER   BY RoutinCd "
     
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        ssSlipList1.Row = ssSlipList1.DataRowCnt + 1
        ssSlipList1.Col = 1
        ssSlipList1.CellType = CellTypeCheckBox
        ssSlipList1.TypeCheckCenter = True
        ssSlipList1.BackColor = RGB(235, 245, 235)
        ssSlipList1.Lock = False
        ssSlipList1.Col = 3: ssSlipList1.ForeColor = RGB(255, 0, 255):
                             ssSlipList1.Text = adoSet.Fields("RoutinCd").Value & ""
        ssSlipList1.Col = 2: ssSlipList1.ForeColor = RGB(255, 0, 255):
                             ssSlipList1.Text = adoSet.Fields("RoutinNm").Value & ""
        ssSlipList1.Col = 4: ssSlipList1.Text = "R"
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return
    
    
SELECT_ItemCode:
    
    strSql = ""
    strSql = strSql & " SELECT Codeky, Itemnm, GbRoutine, GbCheck, GbInput, SUGACD "
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_itemML "
    strSql = strSql & " WHERE  Codeky  LIKE  '" & txtSLipno1.Text & "%'"
    strSql = strSql & " ORDER  BY  GBinput, Codeky"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        ssSlipList1.Row = ssSlipList1.DataRowCnt + 1
        If Trim(adoSet.Fields("gBINPUT").Value & "") = "I" Then
            ssSlipList1.Col = 1
            ssSlipList1.CellType = CellTypeCheckBox
            ssSlipList1.TypeCheckCenter = True
            ssSlipList1.BackColor = RGB(235, 245, 235)
            ssSlipList1.Lock = False
        Else
            ssSlipList1.Col = 1
            ssSlipList1.CellType = SS_CELL_TYPE_EDIT
            ssSlipList1.BackColor = RGB(235, 245, 235)
            ssSlipList1.Text = ""
            ssSlipList1.Lock = True
        End If
        ssSlipList1.Col = 2: ssSlipList1.Text = adoSet.Fields("Itemnm").Value & ""
        ssSlipList1.Col = 3: ssSlipList1.Text = adoSet.Fields("Codeky").Value & ""
        ssSlipList1.Col = 6: ssSlipList1.Text = adoSet.Fields("GbInput").Value & ""
        ssSlipList1.Col = 4: ssSlipList1.Text = adoSet.Fields("GbRoutine").Value & ""
        ssSlipList1.Col = 5: ssSlipList1.Text = adoSet.Fields("GbCheck").Value & ""
        ssSlipList1.Col = 8: ssSlipList1.Text = adoSet.Fields("SUGACD").Value & ""
        
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return

End Sub

Private Sub cmdEnrolOk_Click()
    
    Dim sJdate               As String
    Dim sOrderDt             As String
    Dim sGeomsaDt            As String
    Dim sGeomsaT1            As String
    Dim sGeomsaT2            As String
    Dim nOrderno             As Integer
    Dim iGeneralLabno2       As Integer
    Dim adoSLno2             As ADODB.Recordset
    Dim sRowID               As String
    Dim nSerialNo            As Integer
    
    
    
    GoSub Invalid_Check_Sub
    GoSub ISPatient_Check
    
    sJdate = Format(dtJeobsuDt.Value, "yyyy-MM-dd")
    sOrderDt = Format(dtOrderDt.Value, "yyyy-MM-dd")
    sGeomsaDt = Dual_Date_Get("yyyy-MM-dd")
    sGeomsaT1 = Dual_Date_Get("hh24")
    sGeomsaT2 = Dual_Date_Get("mi")
    
    strSql = " SELECT  TW_MIS_OCS.SEQ_OrderNo.Nextval OdrNo FROM DUAL "
    If adoSetOpen(strSql, adoSet) Then
        nOrderno = Val(adoSet.Fields("OdrNo").Value & "")
        Call adoSetClose(adoSet)
    End If
    
    Call cmdGetno2_Click         'Lab no Setting
    GoSub Serial_PtnoPlusone
    GoSub Idnomst_Process_Sub
    GoSub General_Process_Sub
    GoSub GeneralSub_Process_sub
    
    
    
    
    GLabelJeobsuDt = Format(dtJeobsuDt, "yyyy-MM-dd")
    GLabelPtno = txtPtno.Text
    frmBarCode.Show vbModal
    
    If vbNo = MsgBox("접수되었습니다. 화면을 정리하시겠습니까?", _
                       vbYesNo + vbQuestion, _
                      "접수입력확인Box") Then Exit Sub
    
    Call cmdClear_Click
    
    Exit Sub
    
'/------------------------------------------------------------------------
Invalid_Check_Sub:
    If Trim(txtSLipno1.Text) = "" Then
        MsgBox "입력할 SLip 이 지정되지 않았습니다!.."
        Exit Sub
    End If
    
    If Trim(txtPtno.Text) = "" Then
        MsgBox "입력할 Data(환자)가 선택되지 않았습니다!.."
        Exit Sub
    End If
    
    If Trim(txtSampleCode.Text) = "" Then
        MsgBox "검체코드가 없습니다. 확인하세요!.."
        Exit Sub
    End If
    
    If ssSelectitem.DataRowCnt = 0 Then
        MsgBox "입력할 Data 가 하나도 없습니다!.."
        Exit Sub
    End If
        
    Return
    
    
ISPatient_Check:
    strSql = ""
    strSql = strSql & " SELECT Ptno"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWBAS_PATIENT"
    strSql = strSql & " WHERE  Ptno = '" & txtPtno.Text & "'"
    If False = adoSetOpen(strSql, adoSet) Then
        MsgBox "등록되지 않은 등록번호 입니다!........", vbCritical
        Call adoSetClose(adoSet)
        Exit Sub
    End If
    Call adoSetClose(adoSet)
    Return
    
Serial_PtnoPlusone:
    Dim adoSerial       As ADODB.Recordset
    Dim sSrPtno         As String
    Dim sSrJdate        As String
    
    sSrJdate = Format(dtJeobsuDt.Value, "yyyy-MM-dd")
    sSrPtno = txtPtno.Text
    
    strSql = ""
    strSql = strSql & " SELECT MAX(NVL(dayseq, 0 ) + 1) MaxSerial"
    strSql = strSql & " FROM   TWEXAM_GENERAL_SUB"
    strSql = strSql & " WHERE  Jeobsudt = TO_DATE('" & sSrJdate & "','YYYY-MM-DD')"
    strSql = strSql & " AND    Ptno     = '" & sSrPtno & "'"
    If False = adoSetOpen(strSql, adoSerial) Then
        nSerialNo = 0
    End If
    
    nSerialNo = Val(adoSerial.Fields("MaxSerial").Value & "")
    Call adoSetClose(adoSerial)
    Return
    
    
Idnomst_Process_Sub:
    strSql = ""
    strSql = strSql & " SELECT  * "
    strSql = strSql & " FROM    TWEXAM_IDNOMST "
    strSql = strSql & " WHERE   PtNo =  '" & txtPtno.Text & "'   "

    
    If False = adoSetOpen(strSql, adoSet) Then
        GoSub IDNOMST_INSERT
    Else
        Call adoSetClose(adoSet)
        GoSub IDNOMST_UPDATE
    End If
    Return
    

IDNOMST_INSERT:
    strSql = ""
    strSql = strSql & " INSERT "
    strSql = strSql & " INTO   TWEXAM_IDNOMST"
    strSql = strSql & "       (  PtNo,        Sname,        Sex,         AgeYY, "
    strSql = strSql & "          AgeMM,       Indate,       DeptCode,    RoomCode , "
    strSql = strSql & "          DrCode,      Gbio,         Bi ) "
    strSql = strSql & " VALUES( '" & txtPtno.Text & "',"
    strSql = strSql & "         '" & txtSname.Text & "',"
    strSql = strSql & "         '" & txtSex.Text & "',"
    strSql = strSql & "          " & Val(txtAgeYY.Text) & ","
    strSql = strSql & "          " & Val(txtAgemm.Text) & ","
    strSql = strSql & "              TO_DATE('" & txtIndate.Text & "','YYYY-MM-DD'),"
    strSql = strSql & "         '" & txtDept.Text & "',"
    strSql = strSql & "         '" & txtRoom.Text & "',"
    strSql = strSql & "         '" & txtDr.Text & "',"
    strSql = strSql & "         '" & txtioGubun.Text & "',"
    strSql = strSql & "         '')"
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    Return

'----------------------------------------------------------------------------
IDNOMST_UPDATE:
    strSql = ""
    strSql = strSql & " UPDATE TWEXAM_IDNOMST   "
    strSql = strSql & " SET    Sex      = '" & txtSex.Text & "',"
    strSql = strSql & "        AgeYY    =  " & Val(txtAgeYY.Text) & ","
    strSql = strSql & "        AgeMM    =  " & Val(txtAgemm.Text) & ","
    strSql = strSql & "        Indate   =    TO_DATE('" & txtIndate.Text & "','YYYY-MM-DD'),"
    strSql = strSql & "        DeptCode = '" & txtDept.Text & "',"
    strSql = strSql & "        RoomCode = '" & txtRoom.Text & "',"
    strSql = strSql & "        Gbio     = '" & txtioGubun.Text & "'"
    strSql = strSql & " WHERE  Ptno     = '" & txtPtno.Text & "'"
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    Return


General_Process_Sub:
    
    strSql = ""
    strSql = strSql & " INSERT INTO TWEXAM_GENERAL "
    strSql = strSql & "       (  Jeobsudt,    SlipNo1,      SLipno2,     Jeobsut1, "
    strSql = strSql & "          Jeobsut2,    Jeobsuja,     PtNo,        Codeky,   "
    strSql = strSql & "          Geomchcd,    Geomsagu,     Orderdt,     Orderno,  "
    strSql = strSql & "          Cmdoctor,    RoomCode,     DeptCode,    Gbio,     "
    strSql = strSql & "          DrCode,      inDate,                              "
    strSql = strSql & "          Geomsadt,    Geomsat1,     Geomsat2,              "
    strSql = strSql & "          Geomsaja,    Geomsast,     Geomsacm,    Status,   "
    strSql = strSql & "          Report1,     Sex,          AgeYY,       AgeMM,    "
    strSql = strSql & "          Reporcd,     gbCH )   "
    strSql = strSql & " VALUES(     TO_DATE( '" & sJdate & "','YYYY-MM-DD'),"
    strSql = strSql & "         " & Val(txtSLipno1.Text) & ","
    strSql = strSql & "         " & Val(txtSLipno2.Text) & ","
    strSql = strSql & "         " & Val(Format(dtJeobsuT.Hour, "00")) & ","
    strSql = strSql & "         " & Val(Format(dtJeobsuT.Minute, "00")) & ","
    strSql = strSql & "        '" & txtUserid.Text & "',"
    strSql = strSql & "        '" & txtPtno.Text & "',"
    strSql = strSql & "        '" & txtSLipno1.Text & "',"
    strSql = strSql & "        '" & txtSampleCode.Text & "',"
    strSql = strSql & "        ' ',"
    strSql = strSql & "             TO_DATE('" & sOrderDt & "','YYYY-MM-DD'),"
    strSql = strSql & "         " & nOrderno & ","
    strSql = strSql & "        '" & txtDr.Text & "',"
    strSql = strSql & "        ' ',"
    strSql = strSql & "        '" & txtDept.Text & "',"
    strSql = strSql & "        '" & txtioGubun.Text & "',"
    strSql = strSql & "        '" & txtDr.Text & "',"
    strSql = strSql & "             TO_DATE('" & txtIndate.Text & "','YYYY-MM-DD'),"
    strSql = strSql & "             TO_DATE('" & sGeomsaDt & "','YYYY-MM-DD'),"
    strSql = strSql & "         " & Val(sGeomsaT1) & ","
    strSql = strSql & "         " & Val(sGeomsaT2) & ","
    strSql = strSql & "        '" & txtUserid.Text & "',"
    strSql = strSql & "        ' ',"
    strSql = strSql & "        ' ',"
    strSql = strSql & "        'R',"
    strSql = strSql & "         1, "
    strSql = strSql & "        '" & txtSex.Text & "',"
    strSql = strSql & "         " & Val(txtAgeYY.Text) & ","
    strSql = strSql & "         " & Val(txtAgemm.Text) & ","
    strSql = strSql & "        '" & chkEx.Tag & "',"
    strSql = strSql & "        'Y')"
    
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If
    Return


GeneralSub_Process_sub:
    Dim sInsertItem     As String
    Dim sInsertRtn      As String
    
    For i = 1 To ssSelectitem.DataRowCnt
        ssSelectitem.Row = i
        ssSelectitem.Col = 1
        If ssSelectitem.Value = True Then
            ssSelectitem.Col = 4: sInsertItem = ssSelectitem.Text
            ssSelectitem.Col = 5:
            ssSelectitem.Col = 9: sInsertRtn = ssSelectitem.Text
            GoSub GENERAL_SUB_INSERT_RTN
        End If
    Next
    
    
    Return
    

GENERAL_SUB_INSERT_RTN:
    strSql = ""
    strSql = strSql & " INSERT "
    strSql = strSql & " INTO   TWEXAM_GENERAL_SUB"
    strSql = strSql & "       (JeobsuDt,   SLipno1,   SLipno2,   RoutinCd,   ItemCd, "
    strSql = strSql & "        Ptno,       Sex,       AgeYY,     Agemm,      Orderno,"
    strSql = strSql & "        Verify,     Bi,        GbHost,    GbJeobsu, "
    strSql = strSql & "        Result1,    Result2,   Result3,   Result4,    Result5,"
    strSql = strSql & "        Rcode1,     Rcode2,    Rcode3,    Rcode4,     Rcode5,   Chamgo,"
    strSql = strSql & "        DaySeq)"
    strSql = strSql & " VALUES(      TO_DATE('" & sJdate & "','YYYY-MM-DD'),"
    strSql = strSql & "          " & Val(txtSLipno1.Text) & ","
    strSql = strSql & "          " & Val(txtSLipno2.Text) & ","
    strSql = strSql & "         '" & sInsertRtn & "',"
    strSql = strSql & "         '" & sInsertItem & "',"
    strSql = strSql & "         '" & txtPtno.Text & "',"
    strSql = strSql & "         '" & txtSex.Text & "',"
    strSql = strSql & "          " & Val(txtAgeYY.Text) & ","
    strSql = strSql & "          " & Val(txtAgemm.Text) & ","
    strSql = strSql & "          " & nOrderno & ","
    strSql = strSql & "         'N',"
    strSql = strSql & "         ' ',"
    strSql = strSql & "         '0',"
    strSql = strSql & "         ' ',"
    strSql = strSql & "         ' ',' ',' ',' ',' ',"
    strSql = strSql & "         ' ',' ',' ',' ',' ',"
    strSql = strSql & "         ' ',"
    strSql = strSql & "         " & nSerialNo & ")"
    
    adoConnect.BeginTrans
    If adoExec(strSql) Then
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
    End If

    Return
    

End Sub

Private Sub cmdGetno2_Click()
    Dim sLabelDate      As String
    
    sLabelDate = Format(dtJeobsuDt.Value, "yyyy-MM-dd")
    If cmbioGubun.ListIndex = -1 Then
        MsgBox "입원, 외래 구분을 하십시오!.......", vbCritical
        Exit Sub
    End If
    
    txtSLipno2.Text = Get_Data_Labno(sLabelDate, Val(txtSLipno1.Text), Left(cmbioGubun.Text, 1))
    txtSLipno2.Text = Format(txtSLipno2.Text, "00000")
    

End Sub

Private Sub cmdidClear_Click()

    txtPtno.Text = ""
    txtSname.Text = ""
    txtSex.Text = ""
    txtAgeYY.Text = ""
    txtAgemm.Text = ""
    cmbDept.ListIndex = -1
    txtDept.Text = ""
    cmbDr.ListIndex = -1
    txtDr.Text = ""
    cmbioGubun.ListIndex = -1
    txtioGubun.Text = ""
    txtRoom.Text = ""
    
    chkEx.Tag = ""
    chkEx.Value = "0"
    
    dtOrderDt.Value = Dual_Date_Get("yyyy-MM-dd")
    txtOrderno.Text = ""
    txtIndate.Text = ""
    txtSampleCode.Text = ""
    
    dtJeobsuT = Dual_Date_Get("hh24:mi")
    
    txtUserid = GstrIdnumber
    txtUsername = GstrPassName

End Sub



Private Sub cmdMove_Click()
    Dim sItemCd     As String
    Dim sGbRoutine  As String
    Dim sGbCheck    As String
    Dim sGbInput    As String
    Dim sSugaCd     As String
    
    
    'ssSLipList1 Column
    '/ 1. CheckBox  2.iTemnm  3.Codeky  4.GbRoutine   5.GBCheck  6.Gbinput  8.SugaCD
    Call Spread_Set_Clear(ssSelectitem)
    
    
    For i = 1 To ssSlipList1.DataRowCnt
        ssSlipList1.Row = i
        
        ssSlipList1.Col = 3: sItemCd = ssSlipList1.Text
        ssSlipList1.Col = 4: sGbRoutine = ssSlipList1.Text
        ssSlipList1.Col = 5: sGbCheck = ssSlipList1.Text
        ssSlipList1.Col = 6: sGbInput = ssSlipList1.Text
        ssSlipList1.Col = 8: sSugaCd = ssSlipList1.Text
        ssSlipList1.Col = 1
        If ssSlipList1.CellType = CellTypeCheckBox Then
            If ssSlipList1.Value = True Then
                If sGbRoutine = "R" Then
                    GoSub Get_RoutineCode_Sub
                Else
                    GoSub Get_ItemCode_Sub
                End If
            End If
        End If
    Next
    Exit Sub
    
Get_RoutineCode_Sub:
    strSql = ""
    strSql = strSql & " SELECT ROUTINCD, CODEKY, ITEMNM, SUGACD"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Routine"
    strSql = strSql & " WHERE  ROUTINCD = '" & sItemCd & "'"
    strSql = strSql & " ORDER  BY Codeky"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        ssSelectitem.Row = ssSelectitem.DataRowCnt + 1
        ssSelectitem.Col = 1: ssSelectitem.Value = True
        ssSelectitem.Col = 2: ssSelectitem.Text = adoSet.Fields("itemNm").Value & ""
        ssSelectitem.Col = 4: ssSelectitem.Text = adoSet.Fields("Codeky").Value & ""
        ssSelectitem.Col = 5: ssSelectitem.Text = "R"
        ssSelectitem.Col = 9: ssSelectitem.Text = adoSet.Fields("RoutinCD").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return
    
Get_ItemCode_Sub:
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_itemML"
    strSql = strSql & " WHERE  CODEKY = '" & sItemCd & "'"
    strSql = strSql & " ORDER  BY Codeky"
    If False = adoSetOpen(strSql, adoSet) Then Return
    Do Until adoSet.EOF
        ssSelectitem.Row = ssSelectitem.DataRowCnt + 1
        ssSelectitem.Col = 1: ssSelectitem.Value = True
        ssSelectitem.Col = 2: ssSelectitem.Text = adoSet.Fields("ItemNm").Value & ""
        ssSelectitem.Col = 4: ssSelectitem.Text = adoSet.Fields("Codeky").Value & ""
        ssSelectitem.Col = 5: ssSelectitem.Text = "I"
        ssSelectitem.Col = 9: ssSelectitem.Text = adoSet.Fields("Codeky").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return

End Sub


Private Sub cmdSample_Click()
    
    
    hWndReturn = txtSampleCode.hwnd
    frmQryGeom.Show vbModal
        
    
    strSql = " SELECT Codenm FROM TW_MIS_EXAM.TWEXAM_Sample WHERE CODE = '" & txtSampleCode.Text & "'"
    If False = adoSetOpen(strSql, adoSet) Then
        txtSampleCode.ToolTipText = ""
        Exit Sub
    End If
    
    txtSampleCode.ToolTipText = Trim(adoSet.Fields("Codenm").Value & "")
    Call adoSetClose(adoSet)
    
    
    
End Sub


Private Sub Form_Load()
    
    dtJeobsuDt.Value = Dual_Date_Get("yyyy-MM-dd")
    
    
    GoSub frmManual_Clear_Sub
    GoSub SLip_Select
    GoSub Dept_Setting
    
    Exit Sub
    
frmManual_Clear_Sub:
    For i = 0 To Me.Count - 1
        If TypeOf Me.Controls(i) Is VB.TextBox Then
            Me.Controls(i).Text = ""
        ElseIf TypeOf Me.Controls(i) Is VB.ComboBox Then
            Me.Controls(i).ListIndex = -1
        End If
    Next
    Return


    
    
SLip_Select:
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Specode"
    strSql = strSql & " WHERE  CODEGU = '12'"
    strSql = strSql & " Order  By Codeky"
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        cmbSLip.AddItem Trim(adoSet.Fields("Codeky").Value & "") & ". " & _
                        Trim(adoSet.Fields("Codenm").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return
    
Dept_Setting:
    Dim sDeptCode   As String * 4
    
    strSql = ""
    strSql = strSql & " SELECT *"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWBAS_DEPT"
    strSql = strSql & " WHERE  PrintRanking < 37"
    strSql = strSql & " ORDER  BY PrintRanking"
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sDeptCode = adoSet.Fields("DeptCode").Value & ""
        cmbDept.AddItem sDeptCode & ". " & Trim(adoSet.Fields("DeptnameK").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub ssSlipList1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
    If Row = 0 Then Exit Sub
    
    If Col = 1 Then
        Call cmdMove_Click
    End If
    
End Sub

Private Sub txtPtno_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        txtPtno.Text = UCase(txtPtno.Text)
        txtPtno.Text = Format(txtPtno.Text, "00000000")
        If False = IsAdmission(txtPtno.Text) Then       '입원환자가 아닐때...
            If isPatientMaster(txtPtno.Text) Then       'TW_MIS_PMPA.TWBAS_PATIENT 에 Data 가 있을경우.
                GoSub Patient_Data_Process
            Else
                GoSub IdnoMST_Data_Process
            End If
        Else
            GoSub IPD_Master_Data_Proces               '입원환자일때...
        End If
    End If
    
    
    Exit Sub
    
'/______________________________________________________________
    
IPD_Master_Data_Proces:

    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TWIpd_Master INDEX_IPDMST2) */"
    
    strSql = ""
    strSql = strSql & " SELECT a.*, TO_Char(a.Indate,'YYYY-MM-DD') Indate,"
    strSql = strSql & "        b.Deptnamek, c.Drname"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWIpd_Master a,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT   b,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR c "
    strSql = strSql & " WHERE  a.Ptno     =  '" & txtPtno.Text & "'"
    strSql = strSql & " AND    a.DeptCode = b.Deptcode(+)"
    strSql = strSql & " AND    a.Drcode   = c.Drcode(+)"
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    txtSname.Text = adoSet.Fields("Sname").Value & ""
    txtSex.Text = adoSet.Fields("Sex").Value & ""
    txtAgeYY.Text = adoSet.Fields("Age").Value & ""
    txtDept.Text = adoSet.Fields("DeptCode").Value & ""
    
    Call SetComboBox(cmbDept, adoSet.Fields("DeptCode").Value & "", 4)
    txtDr.Text = adoSet.Fields("Drcode").Value & ""
    
    Call SetComboBox(cmbDr, adoSet.Fields("Drcode").Value & "", 6)
    
    cmbioGubun.ListIndex = 1
    Call cmbioGubun_Click
    txtRoom.Text = adoSet.Fields("RoomCode").Value & ""


    dtOrderDt.Value = dtJeobsuDt.Value
    txtIndate.Text = adoSet.Fields("Indate").Value & ""
    txtUserid.Text = GstrIdnumber
    txtUsername.Text = GstrPassName
    dtJeobsuT.Value = Dual_Date_Get("hh24:mi")
    Return

'/______________________________________________________________

Patient_Data_Process:
    Dim sCompDept   As String * 4
    Dim adoDr       As ADODB.Recordset
    
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INDEX_PATIENT0) */"
    
    strSql = ""
    strSql = strSql & " SELECT a.*, "
    strSql = strSql & "        TO_CHAR(a.LastDate, 'YYYY-MM-DD') LastDate,"
    strSql = strSql & "        b.Drname"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWBAS_PATIENT a,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR  b "
    strSql = strSql & " WHERE  a.Ptno   = '" & txtPtno.Text & "'"
    strSql = strSql & " AND    a.Drcode = b.Drcode(+)"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    txtSname.Text = Trim(adoSet.Fields("Sname").Value & "")
    txtSex.Text = Trim(adoSet.Fields("Sex").Value & "")
    txtAgeYY.Text = SetAge_Check(adoSet.Fields("Jumin1").Value & "", _
                                 adoSet.Fields("Jumin2").Value & "")
    txtAgemm.Text = ""
    sCompDept = adoSet.Fields("DeptCode").Value & ""
    Call SetComboBox(cmbDept, sCompDept, 4)
    txtDept.Text = sCompDept
    
    txtDr.Text = adoSet.Fields("Drcode").Value & ""
    Call SetComboBox(cmbDr, adoSet.Fields("Drcode").Value & "", 6)
    
    cmbioGubun.ListIndex = 1
    Call cmbioGubun_Click
    txtRoom.Text = ""
    
    dtOrderDt.Value = dtJeobsuDt.Value
    txtIndate.Text = adoSet.Fields("LastDate").Value & ""
    txtUserid.Text = GstrIdnumber
    txtUsername.Text = GstrPassName
    dtJeobsuT.Value = Dual_Date_Get("hh24:mi")
    
    Call adoSetClose(adoSet)
    
    Return
    
'/______________________________________________________________
IdnoMST_Data_Process:
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_DEPT INDEX_DEPT0) */"
    
    strSql = ""
    strSql = strSql & " SELECT a.*, TO_Char(a.Indate,'YYYY-MM-DD') Indate,"
    strSql = strSql & "        b.Deptnamek, c.Drname"
    strSql = strSql & " FROM   TWExam_Idnomst a,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT     b,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR   c "
    strSql = strSql & " WHERE  a.Ptno     =  '" & txtPtno.Text & "'"
    strSql = strSql & " AND    a.DeptCode = b.Deptcode(+)"
    strSql = strSql & " AND    a.Drcode   = c.Drcode(+)"
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    txtSname.Text = adoSet.Fields("Sname").Value & ""
    txtSex.Text = adoSet.Fields("Sex").Value & ""
    txtAgeYY.Text = adoSet.Fields("Ageyy").Value & ""
    txtAgemm.Text = adoSet.Fields("Agemm").Value & ""
    txtDept.Text = adoSet.Fields("DeptCode").Value & ""
    
    Call SetComboBox(cmbDept, adoSet.Fields("DeptCode").Value & "", 4)
    txtDr.Text = adoSet.Fields("Drcode").Value & ""
    Call SetComboBox(cmbDr, adoSet.Fields("Drcode").Value & "", 6)
    
    
    txtRoom.Text = adoSet.Fields("RoomCode").Value & ""
    If Trim(txtRoom.Text) = "" Then
        cmbioGubun.ListIndex = 1
    Else
        cmbioGubun.ListIndex = 0
    End If
    Call cmbioGubun_Click

    dtOrderDt.Value = dtJeobsuDt.Value
    txtIndate.Text = adoSet.Fields("Indate").Value & ""
    txtUserid.Text = GstrIdnumber
    txtUsername.Text = GstrPassName
    dtJeobsuT.Value = Dual_Date_Get("hh24:mi")
    
    Return
    

End Sub

Private Sub txtPtno_LostFocus()
    
    
    If Trim(txtPtno.Text) = "" Then Exit Sub
    If False = IsNumeric(txtPtno.Text) Then Exit Sub
    
    txtPtno.Text = Format(txtPtno.Text, "00000000")
    
    
End Sub
