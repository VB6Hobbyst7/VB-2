VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmBarno 
   BackColor       =   &H80000010&
   Caption         =   "검체번호별 Barcode 재발행화면"
   ClientHeight    =   5985
   ClientLeft      =   555
   ClientTop       =   1920
   ClientWidth     =   10995
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
   ScaleHeight     =   5985
   ScaleWidth      =   10995
   WindowState     =   2  '최대화
   Begin Threed.SSPanel SSPanel1 
      Height          =   4965
      Left            =   765
      TabIndex        =   10
      Top             =   765
      Width           =   2940
      _Version        =   65536
      _ExtentX        =   5186
      _ExtentY        =   8758
      _StockProps     =   15
      Caption         =   "SSPanel1"
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
      BorderWidth     =   1
      BevelInner      =   1
      Begin VB.TextBox txtQryPtno 
         Height          =   315
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   0
         Top             =   540
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtJeobsuDt 
         Height          =   330
         Left            =   1080
         TabIndex        =   11
         Top             =   135
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24510467
         CurrentDate     =   36440
      End
      Begin FPSpreadADO.fpSpread sprGeneral 
         Height          =   3660
         Left            =   90
         TabIndex        =   12
         Top             =   1080
         Width           =   2715
         _Version        =   196608
         _ExtentX        =   4789
         _ExtentY        =   6456
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
         MaxCols         =   3
         MaxRows         =   50
         ScrollBars      =   2
         SpreadDesigner  =   "frmBarno.frx":0000
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Caption         =   "접수일자:"
         Height          =   240
         Left            =   135
         TabIndex        =   14
         Top             =   180
         Width           =   870
      End
      Begin VB.Label Label2 
         Caption         =   "등록번호:"
         Height          =   195
         Left            =   135
         TabIndex        =   13
         Top             =   585
         Width           =   870
      End
   End
   Begin VB.TextBox txtPtno 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3825
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "Ptno"
      Top             =   720
      Width           =   1185
   End
   Begin VB.TextBox txtSname 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3825
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "Sname"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtSex 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5265
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "Sex"
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtAge 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5625
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Age"
      Top             =   1080
      Width           =   555
   End
   Begin VB.TextBox txtRoom 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3825
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "RoomCode"
      Top             =   1485
      Width           =   915
   End
   Begin VB.TextBox txtDrcode 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4725
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "Drcode"
      Top             =   1485
      Width           =   915
   End
   Begin VB.TextBox txtDrname 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5625
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "txtDrname"
      Top             =   1485
      Width           =   915
   End
   Begin FPSpreadADO.fpSpread ssLabel 
      Height          =   2895
      Left            =   3780
      TabIndex        =   1
      Top             =   2835
      Width           =   7395
      _Version        =   196608
      _ExtentX        =   13044
      _ExtentY        =   5106
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
      MaxCols         =   16
      ScrollBars      =   2
      SpreadDesigner  =   "frmBarno.frx":0843
      Appearance      =   2
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   8550
      Top             =   855
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   1
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   375
      Left            =   765
      TabIndex        =   17
      Top             =   360
      Width           =   2940
      _Version        =   65536
      _ExtentX        =   5186
      _ExtentY        =   661
      _StockProps     =   15
      Caption         =   "BarCode 재발행"
      ForeColor       =   65535
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "궁서체"
         Size            =   11.99
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSForms.CommandButton cmdPrintOk 
      Height          =   465
      Left            =   8415
      TabIndex        =   16
      Top             =   2295
      Width           =   2715
      Caption         =   "Print"
      PicturePosition =   327683
      Size            =   "4789;820"
      Picture         =   "frmBarno.frx":46E4
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdClear 
      Height          =   465
      Left            =   6120
      TabIndex        =   15
      Top             =   2295
      Width           =   2220
      Caption         =   "Clear"
      PicturePosition =   327683
      Size            =   "3916;820"
      Picture         =   "frmBarno.frx":6E96
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin VB.Line Line1 
      X1              =   3870
      X2              =   10305
      Y1              =   1980
      Y2              =   1980
   End
   Begin MSForms.CommandButton cmdExecute 
      Height          =   465
      Left            =   3780
      TabIndex        =   9
      Top             =   2295
      Width           =   2265
      Caption         =   "조회확인"
      PicturePosition =   327683
      Size            =   "3995;820"
      Picture         =   "frmBarno.frx":8628
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
Attribute VB_Name = "frmBarno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    
    
    
    ssLabel.Row = 1
    ssLabel.Row2 = ssLabel.DataRowCnt + 1
    ssLabel.Col = 1
    ssLabel.Col2 = ssLabel.DataColCnt
    ssLabel.BlockMode = True
    ssLabel.Action = ActionClear
    ssLabel.BlockMode = False

End Sub

Private Sub cmdExecute_Click()
    
    Call Spread_Set_Clear(ssLabel)
    
    txtPtno.Text = GLabelPtno
    GoSub Get_PatientData       '환자정보 Select
    
    GoSub MainProcessing
    
    GoSub ReSelect_Variable
    GoSub Display_ArrayTo_Spread
    Exit Sub
    
    
    
Get_PatientData:
    If IsAdmission(txtPtno.Text) Then
        GoSub Get_ADMaster
    Else
        GoSub Get_HJMaster
    End If
    Return
    


Get_ADMaster:
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TWIPD_MASTER  INDEX_IPDMST2)   "
    'strSql = strSql & "            INDEX (TW_MIS_PMPA.TWBAS_DOCTOR  INDEX_DOCTOR0) */"
    
    strSql = ""
    strSql = strSql & " SELECT a.Sname, a.Sex, a.Age, a.RoomCode, a.DrCode, b.Drname"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWIPD_MASTER a,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR b "
    strSql = strSql & " WHERE  a.Ptno   =  '" & txtPtno.Text & "'"
    strSql = strSql & " AND    a.Drcode = b.Drcode(+)"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    txtSname.Text = adoSet.Fields("Sname").Value & ""
    txtSex.Text = adoSet.Fields("Sex").Value & ""
    txtAge.Text = adoSet.Fields("Age").Value & ""
    txtRoom.Text = adoSet.Fields("RoomCode").Value & ""
    txtDrcode.Text = adoSet.Fields("Drcode").Value & ""
    txtDrname.Text = adoSet.Fields("Drname").Value & ""
    Call adoSetClose(adoSet)
    Return
    

Get_HJMaster:
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INX_PATIENT0)    "
    'strSql = strSql & "            INDEX (TW_MIS_PMPA.TWBAS_DOCTOR  INDEX_DOCTOR0) */"
    
    strSql = ""
    strSql = strSql & " SELECT a.Sname, a.Sex, a.Jumin1, a.Jumin2, a.Drcode,"
    strSql = strSql & "        b.Drname"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWBAS_PATIENT a,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR  b "
    strSql = strSql & " WHERE  a.Ptno    =  '" & txtPtno.Text & "'"
    strSql = strSql & " AND    a.Drcode  = b.Drcode(+)"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    txtSname.Text = adoSet.Fields("Sname").Value & ""
    txtSex.Text = adoSet.Fields("Sex").Value & ""
    txtRoom.Text = ""
    txtAge.Text = SetAge_Check(adoSet.Fields("Jumin1").Value & "", _
                               adoSet.Fields("Jumin2").Value & "")
    txtDrcode.Text = adoSet.Fields("Drcode").Value & ""
    txtDrname.Text = adoSet.Fields("Drname").Value & ""
    Call adoSetClose(adoSet)
    Return
    


MainProcessing:
    'Routine Code 의 약어를 읽지 않고 ItemCode 의 BarText만으로 BartCodePrinting....
    '연속검사의 BarCode 때문에 ...
    
    Call LabelStringClear
    
    strSql = ""
    strSql = strSql & "   SELECT "
    strSql = strSql & "          a.Ptno, TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt, "
    strSql = strSql & "          a.SLipno1, a.SLipno2, b.BarText, b.ChwhYg, e.GeomchCD, b.GeomsaGb,b.BarGb,"
    strSql = strSql & "          e.GbEr, b.ChUnit,"
    strSql = strSql & "          d.Deptnamek"
    strSql = strSql & "   FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "          TW_MIS_EXAM.TWEXAM_itemML      b,"
    strSql = strSql & "          TWEXAM_General     c,"
    strSql = strSql & "          TW_MIS_PMPA.TWBAS_DEPT         d,"
    strSql = strSql & "          TW_MIS_EXAM.TWEXAM_Order       e "
    strSql = strSql & "   WHERE  a.Ptno     =  '" & GLabelPtno & "'"
    strSql = strSql & "   AND    a.JeobsuDt = TO_DATE('" & GLabelJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & "   AND    a.SLipno1  =  " & GLabelLabno1
    strSql = strSql & "   AND    a.SLipno2  =  " & GLabelLabno2
    strSql = strSql & "   AND   ( a.Routincd = a.Itemcd Or b.barGb = '1')"
    strSql = strSql & "   AND    a.ItemCD   = b.Codeky"
    strSql = strSql & "   AND    a.JeobsuDt = c.JeobsuDt(+)"
    strSql = strSql & "   AND    a.SLipno1  = c.SLipno1(+)"
    strSql = strSql & "   AND    a.SLipno2  = c.SLipno2(+)"
    strSql = strSql & "   AND    c.DeptCode = d.DeptCode(+)"
    strSql = strSql & "   AND    a.JeobsuDt = e.CollDate(+)"
    strSql = strSql & "   AND    a.SLipno1  = e.SLipno1(+)"
    strSql = strSql & "   AND    a.Orderno  = e.Orderno(+)"
    strSql = strSql & "   GROUP BY a.Ptno, a.JeobsuDt, a.SLipno1, a.SLipno2, b.BarText, b.Chwhyg, e.GeomchCD, "
    strSql = strSql & "            b.GeomsaGb, b.BarGb, e.GbEr, b.ChUnit, d.Deptnamek"
    strSql = strSql & "   UNION ALL"
    strSql = strSql & "   SELECT a.Ptno, TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt, "
    strSql = strSql & "          a.SLipno1, a.SLipno2, d.Yakcd BarText, b.ChwhYg, f.GeomchCD, b.GeomsaGb,b.BarGb,"
    strSql = strSql & "          f.GbEr, b.ChUnit,"
    strSql = strSql & "          e.Deptnamek"
    strSql = strSql & "   FROM   TWEXAM_General_Sub a,"
    strSql = strSql & "          TW_MIS_EXAM.TWEXAM_itemML      b,"
    strSql = strSql & "          TWEXAM_General     c,"
    strSql = strSql & "          TW_MIS_EXAM.TWEXAM_Routine     d,"
    strSql = strSql & "          TW_MIS_PMPA.TWBAS_DEPT         e,"
    strSql = strSql & "          TW_MIS_EXAM.TWEXAM_Order       f"
    strSql = strSql & "   WHERE  a.Ptno     =  '" & GLabelPtno & "'"
    strSql = strSql & "   AND    a.JeobsuDt = TO_DATE('" & GLabelJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & "   AND    a.SLipno1  =  " & GLabelLabno1
    strSql = strSql & "   AND    a.SLipno2  =  " & GLabelLabno2
    strSql = strSql & "   AND    a.Routincd != a.Itemcd"
    strSql = strSql & "   AND    a.ItemCD    = b.Codeky"
    strSql = strSql & "   AND   (b.BarGB IS NULL OR b.BarGB != '1')"
    strSql = strSql & "   AND    a.JeobsuDt  = c.JeobsuDt(+)"
    strSql = strSql & "   AND    a.SLipno1   = c.SLipno1(+)"
    strSql = strSql & "   AND    a.SLipno2   = c.SLipno2(+)"
    strSql = strSql & "   AND    a.Routincd  = d.RoutinCd"
    strSql = strSql & "   AND   ( d.Series    IS NULL Or d.Series != '1')"
    strSql = strSql & "   AND    c.DeptCode  = e.DeptCode(+)"
    strSql = strSql & "   AND    a.JeobsuDt  = f.CollDate(+)"
    strSql = strSql & "   AND    a.SLipno1   = f.SLipno1(+)"
    strSql = strSql & "   AND    a.Orderno   = f.Orderno(+)"
    strSql = strSql & "   GROUP BY a.Ptno, a.JeobsuDt, a.SLipno1, a.SLipno2, d.YakCD, b.Chwhyg, f.GeomchCD, "
    strSql = strSql & "            b.GeomsaGb, b.BarGb, f.GbEr, b.ChUnit, e.Deptnamek"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    i = 0
    Do Until adoSet.EOF
        LabelString.Ptno(i) = adoSet.Fields("Ptno").Value & ""
        LabelString.JeobsuDt(i) = adoSet.Fields("JeobsuDt").Value & ""
        LabelString.sLipno1(i) = adoSet.Fields("SLipno1").Value & ""
        LabelString.Slipno2(i) = adoSet.Fields("SLipno2").Value & ""
        LabelString.BarText(i) = adoSet.Fields("BarText").Value & ""
        LabelString.Yg(i) = adoSet.Fields("Chwhyg").Value & ""
        LabelString.SampleCd(i) = adoSet.Fields("GeomchCD").Value & ""
        LabelString.ReporCd(i) = adoSet.Fields("GeomsaGb").Value & ""
        LabelString.Er(i) = adoSet.Fields("GbEr").Value & ""
        LabelString.DeptCode(i) = Trim(adoSet.Fields("DeptNamek").Value & "")
        LabelString.ChUnit(i) = Trim(adoSet.Fields("ChUnit").Value & "")
        
        LabelString.Title(i) = LabelString.Ptno(i) & _
                               LabelString.JeobsuDt(i) & _
                               LabelString.sLipno1(i) & _
                               LabelString.Slipno2(i) & _
                               LabelString.Yg(i) & _
                               LabelString.SampleCd(i) & _
                               LabelString.ReporCd(i) & _
                               LabelString.Er(i)
        
        If adoSet.Fields("BarGB").Value & "" = "1" Then            'BarCode Label 을 따로 관리하는 항목은 ....
            LabelString.Title(i) = LabelString.Title(i) & LabelString.BarText(i)
        End If
                               
        adoSet.MoveNext: i = i + 1
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
    
ReSelect_Variable:
    Dim nStart       As String
    
    Call LabelString1Clear
    
    For i = 0 To 50
        If isArrayText(LabelString1.Title, LabelString.Title(i)) Then
            If Trim(LabelString.BarText(i)) <> "" Then
                LabelString1.BarText(GVarPoint) = LabelString1.BarText(GVarPoint) & "," & LabelString.BarText(i)
            End If
        Else
            nStart = isArrayMaxReturn(LabelString1.Title)
            LabelString1.Title(nStart) = LabelString.Title(i)
            LabelString1.Ptno(nStart) = LabelString.Ptno(i)
            LabelString1.JeobsuDt(nStart) = LabelString.JeobsuDt(i)
            LabelString1.sLipno1(nStart) = LabelString.sLipno1(i)
            LabelString1.Slipno2(nStart) = LabelString.Slipno2(i)
            LabelString1.BarText(nStart) = LabelString.BarText(i)
            LabelString1.Yg(nStart) = LabelString.Yg(i)
            LabelString1.SampleCd(nStart) = LabelString.SampleCd(i)
            LabelString1.ReporCd(nStart) = LabelString.ReporCd(i)
            LabelString1.Er(nStart) = LabelString.Er(i)
            LabelString1.DeptCode(nStart) = LabelString.DeptCode(i)
            LabelString1.ChUnit(nStart) = LabelString.ChUnit(i)
        End If
    Next
    Return


Display_ArrayTo_Spread:
    Call Spread_Set_Clear(ssLabel)
    
    For i = 0 To 50
        ssLabel.Row = i + 1
        If LabelString1.Title(i) <> "" Then
            ssLabel.Col = 1:  ssLabel.Value = True
            ssLabel.Col = 2:  ssLabel.Text = LabelString1.Ptno(i)
            ssLabel.Col = 3:  ssLabel.Text = txtSname.Text
            ssLabel.Col = 4:  ssLabel.Text = txtRoom.Text
                                
            ssLabel.Col = 5:  ssLabel.Text = LabelString1.JeobsuDt(i)
            ssLabel.Col = 6:  ssLabel.Text = LabelString1.sLipno1(i)
            ssLabel.Col = 7:  ssLabel.Text = Format(LabelString1.Slipno2(i), "00000")
            ssLabel.Col = 8:  ssLabel.Text = LabelString1.BarText(i)
            ssLabel.Col = 9:  ssLabel.TypeComboBoxCurSel = Val(LabelString1.ChUnit(i))
            ssLabel.Col = 10: ssLabel.Text = LabelString1.SampleCd(i)
                              GoSub Get_SampleData
            ssLabel.Col = 12: ssLabel.Text = LabelString1.Yg(i)
                              'GoSub Get_YgData
            ssLabel.Col = 14: ssLabel.Text = LabelString1.Er(i)
            ssLabel.Col = 15: ssLabel.Text = LabelString1.ReporCd(i)
            ssLabel.Col = 16: ssLabel.Text = LabelString1.DeptCode(i)
        End If
    Next
    
    Return
    
    
Get_SampleData:
    strSql = ""
    strSql = strSql & " SELECT * "
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Sample"
    strSql = strSql & " WHERE  Code = '" & LabelString1.SampleCd(i) & "'"
    If False = adoSetOpen(strSql, adoSet) Then Return
    ssLabel.Col = 11: ssLabel.Text = Trim(adoSet.Fields("Codenm").Value & "")
    Call adoSetClose(adoSet)
    Return

Get_YgData:
    strSql = ""
    strSql = strSql & " SELECT CODENM, Yageo"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Specode"
    strSql = strSql & " WHERE  CODEGU = '88'"
    strSql = strSql & " AND    CODEKY = '" & LabelString1.Yg(i) & "'"
    If False = adoSetOpen(strSql, adoSet) Then Return
    ssLabel.Col = 13: ssLabel.Text = Trim(adoSet.Fields("Yageo").Value & "")
    Call adoSetClose(adoSet)
    Return

End Sub

Private Sub cmdPrintOk_Click()
    Dim sBarCodeText(8) As String
    
    Dim sBarSLno1       As String
    Dim sBarSLno2       As String
    Dim sBarJdate       As String
    Dim sBarText        As String
    Dim nLoop           As Integer
    Dim sSLName         As String
    Dim sBarRoom        As String
    Dim sEr             As String
    Dim sEx             As String
    Dim sSample         As String
    Dim sDeptCode       As String
    Dim sSLipText       As String
    
    
    If ssLabel.DataRowCnt = 0 Then
        MsgBox "Barcode Printing 할 Data 가 하나도 없습니다!.."
        Exit Sub
    End If
    
    For i = 1 To ssLabel.DataRowCnt
        
        ssLabel.Row = i
        ssLabel.Col = 1
        If ssLabel.Value = True Then
            GoSub Set_Array_Clear
            ssLabel.Col = 16: sDeptCode = Trim(ssLabel.Text)
            ssLabel.Col = 15: sEx = Trim(ssLabel.Text)
            ssLabel.Col = 14: sEr = Trim(ssLabel.Text)
            ssLabel.Col = 11: sSample = Trim(ssLabel.Text)
            ssLabel.Col = 6:  sBarSLno1 = ssLabel.Text
            ssLabel.Col = 7:  sBarSLno2 = Format(ssLabel.Text, "00000")
            ssLabel.Col = 8:  sBarText = ssLabel.Text
            ssLabel.Col = 5:  sBarJdate = ssLabel.Text
            ssLabel.Col = 9:  nLoop = Val(ssLabel.Text)        'Print 장수
            ssLabel.Col = 4:  sBarRoom = Trim(ssLabel.Text)    '병실Code
            
            'GoSub GET_SLipname
            sBarJdate = Replace(sBarJdate, "-", "", 1, , vbTextCompare)
            sSLipText = convSLipYageo(sBarSLno1)
            
            sBarCodeText(0) = sSLipText                        '11:(혈액검사일반) 12:(혈액검사특수) ...
            sBarCodeText(1) = sSample
            
            If Trim(sEr) <> "" Then
                sBarCodeText(2) = "응급": End If               'SBARCODETEXT(2) = "응급/(외),''"
                
            If Trim(sEx) = "W" Then
                If Trim(sBarCodeText(2)) = "" Then
                    sBarCodeText(2) = "(외)"
                Else
                    sBarCodeText(2) = sBarCodeText(2) & "/" & "(외)"
                End If
            End If
            
            'sBarCodeText(2) = "응급/(외)"
            
            sBarCodeText(3) = sBarJdate & "-" & sSLipText & "  " & sBarSLno2
            sBarCodeText(4) = sBarJdate & sBarSLno1 & sBarSLno2
            sBarCodeText(5) = txtPtno.Text
            sBarCodeText(6) = txtSname.Text
            'sBarCodeText(5) = txtPtno.Text & "," & txtSname.Text & "," & txtSex.Text & "/" & txtAge.Text

            If Trim(sBarRoom) = "" Then
                sBarCodeText(6) = sBarCodeText(6) & "," & sDeptCode
            Else
                sBarCodeText(6) = sBarCodeText(6) & "," & sBarRoom
            End If
                
            sBarCodeText(7) = sBarText
'            Call Bar7421_Printing_Sub(sBarCodeText, nLoop, MSComm1)
            Call BarCodePrint(sBarCodeText, nLoop, Me)
        End If
    Next
    Exit Sub
    

    
    
Set_Array_Clear:
    Dim iVar        As Integer
    
    For iVar = 0 To 7
        sBarCodeText(iVar) = ""
    Next
    
    sEx = ""
    sEr = ""
    sBarSLno1 = ""
    sBarSLno2 = ""
    sBarText = ""
    sBarJdate = ""
    nLoop = 0
    sBarRoom = ""
    
    Return
    
    
GET_SLipname:
    strSql = " SELECT Yageo FROM TW_MIS_EXAM.TWEXAM_Specode WHERE CODEGU = '12' AND Codeky = '" & sBarSLno1 & "'"
    If False = adoSetOpen(strSql, adoSet) Then
        sSLName = ""
        Return
    End If
    sSLName = "[" & Trim(adoSet.Fields("Yageo").Value & "") & "]"
    Call adoSetClose(adoSet)
    Return
    

End Sub

Private Sub Form_Activate()
    
    Me.WindowState = vbMaximized
    
End Sub

Private Sub Form_Load()
    
    dtJeobsuDt.Value = Dual_Date_Get("yyyy-MM-dd")
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub sprGeneral_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then Exit Sub
    If Row > sprGeneral.DataRowCnt Then Exit Sub
    
    GLabelPtno = txtQryPtno.Text
    GLabelJeobsuDt = Format(dtJeobsuDt.Value, "yyyy-MM-dd")
    sprGeneral.Row = Row
    sprGeneral.Col = 1: GLabelLabno1 = Val(sprGeneral.Text)
    sprGeneral.Col = 3: GLabelLabno2 = Val(sprGeneral.Text)
    
    sprGeneral.Row = -1
    sprGeneral.Col = -1
    sprGeneral.ForeColor = RGB(0, 0, 0)
    
    sprGeneral.Row = Row
    sprGeneral.Row2 = Row
    sprGeneral.Col = 1
    sprGeneral.Col2 = sprGeneral.MaxCols
    sprGeneral.BlockMode = True
    sprGeneral.ForeColor = RGB(0, 0, 255)
    sprGeneral.BlockMode = False
    
    
    Call cmdExecute_Click
    
    
End Sub

Private Sub txtQryPtno_KeyPress(KeyAscii As Integer)
    Dim sJeobsuDt       As String
    
    
    If KeyAscii = 13 Then
        Call Spread_Set_Clear(sprGeneral)
        sJeobsuDt = Format(dtJeobsuDt.Value, "yyyy-MM-dd")
        txtQryPtno.Text = Format(txtQryPtno.Text, "00000000")
        GoSub Main_Process_Sub
        
        If Me.sprGeneral.DataRowCnt > 0 Then
            Call sprGeneral_DblClick(1, 1)
        End If
    
    End If
    
    
    Exit Sub
    
Main_Process_Sub:
    strSql = ""
    strSql = strSql & " SELECT a.SLipno1, b.Codenm SLipname, a.SLipno2"
    strSql = strSql & " FROM   TWEXAM_General a,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Specode b"
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.Ptno     = '" & txtQryPtno.Text & "'"
    strSql = strSql & " AND    a.SLipno1  = TO_NUMBER(b.Codeky)"
    strSql = strSql & " AND    b.Codegu   = '12'"
    strSql = strSql & " ORDER  BY SLipno1 ASC, SLipno2 DESC"
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sprGeneral.Row = sprGeneral.DataRowCnt + 1
        sprGeneral.Col = 1: sprGeneral.Text = adoSet.Fields("SLipno1").Value & ""
        sprGeneral.Col = 2: sprGeneral.Text = adoSet.Fields("SLipname").Value & ""
        sprGeneral.Col = 3: sprGeneral.Text = adoSet.Fields("SLipno2").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return
    
    
    
End Sub
