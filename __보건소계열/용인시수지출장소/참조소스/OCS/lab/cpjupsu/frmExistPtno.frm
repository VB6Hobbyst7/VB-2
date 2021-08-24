VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmExistPtno 
   Caption         =   "접수환자병록번호별 조회화면"
   ClientHeight    =   7935
   ClientLeft      =   210
   ClientTop       =   915
   ClientWidth     =   11805
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
   ScaleHeight     =   7935
   ScaleWidth      =   11805
   Begin VB.CheckBox chkOPd 
      Caption         =   "Check2"
      Height          =   375
      Left            =   2925
      TabIndex        =   28
      Top             =   2745
      Width           =   1860
   End
   Begin VB.CheckBox chkIPd 
      Height          =   240
      Left            =   1080
      TabIndex        =   27
      Top             =   2790
      Width           =   1545
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   2400
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   11580
      _Version        =   65536
      _ExtentX        =   20426
      _ExtentY        =   4233
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
      BorderWidth     =   2
      Alignment       =   0
      Begin Threed.SSPanel SSPanel2 
         Height          =   1455
         Left            =   135
         TabIndex        =   19
         Top             =   675
         Width           =   2670
         _Version        =   65536
         _ExtentX        =   4710
         _ExtentY        =   2566
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
         Begin VB.TextBox txtEptno 
            Height          =   285
            Left            =   945
            TabIndex        =   25
            Top             =   945
            Width           =   1230
         End
         Begin VB.OptionButton Option2 
            Caption         =   "입원"
            Height          =   195
            Left            =   135
            TabIndex        =   21
            Top             =   360
            Width           =   690
         End
         Begin VB.OptionButton Option1 
            Caption         =   "외래"
            Height          =   240
            Left            =   135
            TabIndex        =   20
            Top             =   90
            Value           =   -1  'True
            Width           =   690
         End
         Begin Threed.SSPanel panelVerify 
            Height          =   510
            Left            =   855
            TabIndex        =   22
            Top             =   360
            Visible         =   0   'False
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   900
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
            BevelOuter      =   0
            BevelInner      =   1
            Begin VB.OptionButton Option3 
               Caption         =   "확인"
               Height          =   180
               Left            =   135
               TabIndex        =   24
               Top             =   45
               Value           =   -1  'True
               Width           =   690
            End
            Begin VB.OptionButton Option4 
               Caption         =   "미확인"
               Height          =   180
               Left            =   135
               TabIndex        =   23
               Top             =   270
               Width           =   870
            End
         End
         Begin VB.Label Label1 
            Caption         =   "등록번호"
            Height          =   195
            Left            =   135
            TabIndex        =   26
            Top             =   990
            Width           =   780
         End
      End
      Begin VB.TextBox txtAddr 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   6075
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   810
         Width           =   3885
      End
      Begin VB.TextBox txtTel 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   6075
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   495
         Width           =   1680
      End
      Begin VB.TextBox txtSname 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   6075
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "홍길동아가"
         Top             =   180
         Width           =   1005
      End
      Begin VB.TextBox txtJumin1 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   8595
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "670815"
         Top             =   180
         Width           =   645
      End
      Begin VB.TextBox txtJumin2 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   9225
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "1462411"
         Top             =   180
         Width           =   735
      End
      Begin VB.TextBox txtBirthDate 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   8595
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "BirthDate"
         Top             =   495
         Width           =   1365
      End
      Begin VB.TextBox txtSex 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   7110
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "M"
         Top             =   180
         Width           =   225
      End
      Begin VB.TextBox txtAgeYY 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "999"
         Top             =   180
         Width           =   420
      End
      Begin MSComCtl2.DTPicker dtFrDate 
         Height          =   285
         Left            =   135
         TabIndex        =   2
         Top             =   315
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24379395
         CurrentDate     =   36413
      End
      Begin MSComCtl2.DTPicker dtToDate 
         Height          =   285
         Left            =   1530
         TabIndex        =   5
         Top             =   315
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24379395
         CurrentDate     =   36413
      End
      Begin VB.Line Line1 
         X1              =   5085
         X2              =   5085
         Y1              =   90
         Y2              =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "환자명"
         Height          =   195
         Left            =   5400
         TabIndex        =   18
         Top             =   225
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "주민번호"
         Height          =   240
         Left            =   7830
         TabIndex        =   17
         Top             =   225
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "주소"
         Height          =   240
         Index           =   5
         Left            =   5580
         TabIndex        =   16
         Top             =   810
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "전화번호"
         Height          =   240
         Index           =   4
         Left            =   5220
         TabIndex        =   15
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "생년월일"
         Height          =   240
         Left            =   7830
         TabIndex        =   14
         Top             =   540
         Width           =   825
      End
      Begin MSForms.CommandButton cmdExistOk 
         Height          =   915
         Left            =   10215
         TabIndex        =   4
         Top             =   180
         Width           =   915
         Caption         =   "조회확인"
         Size            =   "1614;1614"
         FontName        =   "굴림체"
         FontHeight      =   180
         FontCharSet     =   129
         FontPitchAndFamily=   18
         ParagraphAlign  =   3
      End
      Begin VB.Label Label5 
         Caption         =   "검체접수일:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   90
         Width           =   1005
      End
   End
   Begin FPSpreadADO.fpSpread sprOrder 
      Height          =   4470
      Left            =   90
      TabIndex        =   0
      Top             =   3330
      Width           =   11715
      _Version        =   196608
      _ExtentX        =   20664
      _ExtentY        =   7885
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   1
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
      GridShowHoriz   =   0   'False
      MaxCols         =   36
      ScrollBars      =   2
      ShadowColor     =   12632256
      ShadowDark      =   8421504
      ShadowText      =   0
      SpreadDesigner  =   "frmExistPtno.frx":0000
      Appearance      =   1
      TextTip         =   1
      ScrollBarTrack  =   1
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmExistPtno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExistOk_Click()
    Dim sFrDate             As String
    Dim sToDate             As String
    Dim sCompare            As String

    Screen.MousePointer = vbHourglass
    
    
    GoSub Form_Clear_Sub
    
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    GoSub Get_Order_MainProcess
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
Form_Clear_Sub:
    sprOrder.Row = -1
    sprOrder.Col = -1
    sprOrder.MaxRows = 0
    sprOrder.MaxRows = 500
    sprOrder.RowHeight(-1) = 10.5
    
    Return
    
    
Get_Order_MainProcess:

    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INX_PATIENT0)    "
    'strSql = strSql & "            INDEX (TW_MIS_PMPA.TWBAS_DOCTOR  INDEX_DOCTOR0) */"
    
    strSql = ""
    strSql = strSql & " SELECT a.*, a.RowID OrderRowID,"
    strSql = strSql & "        TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') Jeobsudt,"
    strSql = strSql & "        TO_CHAR(a.Indate,   'YYYY-MM-DD') Indate,  "
    strSql = strSql & "        TO_CHAR(a.OrderDt,  'YYYY-MM-DD') Orderdt, "
    strSql = strSql & "        TO_CHAR(a.CollDate, 'YYYY-MM-DD') CollDate,"
    strSql = strSql & "        b.Sname, c.Codenm SLname,"
    strSql = strSql & "        d.Codenm Samplename, e.Drname, f.JeobsuJa"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Order   a, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PATIENT  b, "
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Specode c, "
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Sample  d, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR   e, "
    strSql = strSql & "        TWEXAM_General f  "
    strSql = strSql & " WHERE  a.CollDate  >=   TO_DATE('" & sFrDate & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.CollDate  <=   TO_DATE('" & sToDate & "','YYYY-MM-DD')"
    If Trim(txtEptno.Text) <> "" Then
        strSql = strSql & " AND    a.Ptno       =   '" & txtEptno.Text & "'"
    End If
    strSql = strSql & " AND    a.JeobsuYN  =  '*'"
    'C strSql = strSql & " AND    a.SLipno1  <   52"
    strSql = strSql & " AND    a.SLipno1  <   90"

    If chkOPd.Value = "1" Then
        If chkIPd.Value = "0" Then
            strSql = strSql & " AND    a.Gbio      = 'O'"         '외래환자만
        End If
    Else
        If chkOPd.Value = "0" Then
            strSql = strSql & " AND    a.Gbio      = 'I'"         '입원환자만
        End If
    End If
        
    strSql = strSql & " AND    a.Ptno      = b.Ptno(+)"
    strSql = strSql & " AND    c.Codegu    = '12'"
    strSql = strSql & " AND    a.GeomchCd  = d.Code(+)"
    strSql = strSql & " AND    a.Drcode    = e.Drcode(+)"
    strSql = strSql & " AND    TO_NUMBER(c.Codeky)  = a.SLipno1"
    strSql = strSql & " AND    a.JeobsuDt  = f.JeobsuDt(+)"
    strSql = strSql & " AND    a.SLipno1   = f.SLipno1(+)"
    strSql = strSql & " AND    a.Orderno   = f.Orderno(+)"
    If gSver = "Y" Then
        strSql = strSql & " AND  f.GBCH = 'Y'"
    Else
        strSql = strSql & " AND  f.GBCH IN ('1','2')"
    End If
    
    strSql = strSql & " ORDER  BY a.CollDate, a.Ptno, a.DeptCode, a.SLipno1"
    
    sprOrder.MaxRows = 0
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    sprOrder.MaxRows = adoSet.RecordCount
    
    Do Until adoSet.EOF
        If sprOrder.Row = sprOrder.MaxRows Then
            sprOrder.MaxRows = sprOrder.MaxRows + 1
            sprOrder.RowHeight(sprOrder.MaxRows) = 10.5
        End If
        
        sprOrder.Row = sprOrder.DataRowCnt + 1
        sprOrder.Col = 2:  sprOrder.Text = adoSet.Fields("JeobsuDt").Value & "" & _
                                           adoSet.Fields("Ptno").Value & ""

        sprOrder.Col = 2
        If sCompare <> sprOrder.Text Then
            sprOrder.Col = 4:  sprOrder.Text = Format(adoSet.Fields("collHH").Value, "00") & ":" & _
                                               Format(adoSet.Fields("collMM").Value, "00")
            
            sprOrder.Col = 5:  sprOrder.Text = adoSet.Fields("Ptno").Value & ""
            sprOrder.Col = 6:  sprOrder.Text = adoSet.Fields("Sname").Value & ""
            sprOrder.Col = 7:  sprOrder.Text = adoSet.Fields("Sex").Value & ""
            sprOrder.Col = 8:  sprOrder.Text = adoSet.Fields("AgeYY").Value & ""
            sprOrder.Col = 9:  sprOrder.Text = adoSet.Fields("AgeMM").Value & ""
            sprOrder.Col = 1:  sprOrder.CellType = CellTypeButton
            Call SpreadRowTopLine(sprOrder, sprOrder.Row)
        End If
        
        sprOrder.Col = 3:   sprOrder.Text = adoSet.Fields("OrderRowID").Value & ""
        sprOrder.Col = 10:  sprOrder.Text = adoSet.Fields("SLipno1").Value & ""
        sprOrder.Col = 11:  sprOrder.Text = adoSet.Fields("SLname").Value & ""
                
        sprOrder.Col = 12: sprOrder.Text = adoSet.Fields("Itemcd").Value & ""
        
        If IsRoutineCode(adoSet.Fields("ItemCd").Value & "") Then
            sprOrder.Col = 13: sprOrder.Text = Get_RoutineName(adoSet.Fields("ItemCD").Value & "")
        Else
            sprOrder.Col = 13: sprOrder.Text = Get_ItemName(adoSet.Fields("ItemCD").Value & "")
        End If
                
                
        sprOrder.Col = 14:  sprOrder.Text = Format(adoSet.Fields("collHH").Value, "00") & ":" & _
                                            Format(adoSet.Fields("collMM").Value, "00")
        
        sprOrder.Col = 15: sprOrder.Text = adoSet.Fields("Indate").Value & ""
        sprOrder.Col = 16: sprOrder.Text = adoSet.Fields("RoomCode").Value & ""
        sprOrder.Col = 17: sprOrder.Text = adoSet.Fields("DeptCode").Value & ""
        sprOrder.Col = 18: sprOrder.Text = adoSet.Fields("Gbio").Value & ""
        sprOrder.Col = 19: sprOrder.Text = adoSet.Fields("Bi").Value & ""
        sprOrder.Col = 20: sprOrder.Text = adoSet.Fields("GbER").Value & ""
        sprOrder.Col = 21: sprOrder.Value = True
        
        sprOrder.Col = 22: sprOrder.Text = adoSet.Fields("GeomchCD").Value & ""
        sprOrder.Col = 23: sprOrder.Text = adoSet.Fields("Samplename").Value & ""
        
        sprOrder.Col = 24: sprOrder.Text = adoSet.Fields("GeomsaGu").Value & ""
        sprOrder.Col = 25: sprOrder.Text = adoSet.Fields("OrderDt").Value & ""
        sprOrder.Col = 26: sprOrder.Text = adoSet.Fields("OrderNo").Value & ""
        sprOrder.Col = 27: sprOrder.Text = adoSet.Fields("OrderCD").Value & ""
        sprOrder.Col = 28: sprOrder.Text = adoSet.Fields("Quantity").Value & ""
        sprOrder.Col = 29: sprOrder.Text = adoSet.Fields("CmDoctor").Value & ""
        sprOrder.Col = 30: sprOrder.Text = adoSet.Fields("DrCode").Value & ""
        sprOrder.Col = 31: sprOrder.Text = adoSet.Fields("Drname").Value & ""
        sprOrder.Col = 32: sprOrder.Text = adoSet.Fields("JeobsuYn").Value & ""
        sprOrder.Col = 33: sprOrder.Text = adoSet.Fields("Gbinfo").Value & ""
        
        
        sprOrder.Col = 34: sprOrder.Text = adoSet.Fields("JeobsuDt").Value & ""
        
        sprOrder.Col = 35:  sprOrder.Text = Format(adoSet.Fields("JeobsuT1").Value, "00") & ":" & _
                                            Format(adoSet.Fields("JeobsuT2").Value, "00")
        sprOrder.Col = 36:  sprOrder.Text = adoSet.Fields("JeobsuJa").Value & ""
                            GoSub Get_JeobsuJaName
        sCompare = adoSet.Fields("JeobsuDt").Value & "" & _
                   adoSet.Fields("Ptno").Value & ""
        
        
        adoSet.MoveNext
    Loop
    
    Call adoSetClose(adoSet)
    
    Return

Get_JeobsuJaName:
    Dim sJcode(1)    As String
    Dim adoPass      As ADODB.Recordset
    
    If sJcode(0) = adoSet.Fields("JeobsuJa").Value & "" Then
        sprOrder.Col = 36: sprOrder.Text = sJcode(1)
        Return
    End If
    
    sJcode(0) = adoSet.Fields("JeobsuJa").Value & ""
    strSql = " SELECT NAME FROM  TW_MIS_PMPA.TWBAS_Pass WHERE iDNumber = '" & sJcode(0) & "'"
    If False = adoSetOpen(strSql, adoPass) Then
        Return
    End If
    
    sprOrder.Col = 36:  sprOrder.Text = adoPass.Fields("Name").Value & ""
    sJcode(1) = sprOrder.Text
    
    Call adoSetClose(adoPass)
    
    
    Return

End Sub

Private Sub cmdQryAll_Click()

End Sub

Private Sub Form_Load()
    
    If gSio = "I" Then
        chkIPd.Value = "1"
        chkOPd.Value = "0"
    Else
        chkOPd.Value = "1"
        chkIPd.Value = "0"
    End If
        
    
    GoSub Date_Setting
    Exit Sub
    
Date_Setting:
    dtFrDate.Value = Dual_Date_Get("yyyy-MM-dd")
    dtToDate.Value = Dual_Date_Get("yyyy-MM-dd")
    
    Return
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub Option1_Click()
    
    If Option1.Value = True Then
        panelVerify.Visible = False
    End If

End Sub

Private Sub Option2_Click()
    
    If Option2.Value = True Then
        panelVerify.Visible = True
    End If
    
End Sub

Private Sub sprOrder_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    If Col = 1 Then
        sprOrder.Row = Row
        sprOrder.Col = 34: GLabelJeobsuDt = sprOrder.Text
        sprOrder.Col = 5: GLabelPtno = sprOrder.Text
        frmBarCode.Show vbModal

    End If

End Sub

Private Sub txtEptno_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Trim(txtEptno.Text) = "" Then Exit Sub
        txtEptno.Text = Format(txtEptno.Text, "00000000")
        DoEvents: GoSub Get_Hj_MasterData
        DoEvents: Call cmdExistOk_Click
    End If
    Exit Sub
    
    
Get_Hj_MasterData:
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INX_PATIENT0) "
    'strSql = strSql & "            INDEX (TWBAS_POST    INDEX_POST2) */"
    
    strSql = ""
    strSql = strSql & " SELECT TO_CHAR(a.BirthDate, 'YYYY-MM-DD') BirthDate,"
    strSql = strSql & "        a.Ptno,   a.Sname,"
    strSql = strSql & "        a.Jumin1, a.Jumin2, a.Juso, a.Tel, a.Sex, a.Bi, a.Juso,"
    strSql = strSql & "        b.PostName1, b.PostName2, b.PostName3"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWBAS_PATIENT a,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_POST    b "
    strSql = strSql & " WHERE  a.PTNO      = '" & txtEptno.Text & "'"
    strSql = strSql & " AND    a.PostCode1 = b.PostCode1(+)"
    strSql = strSql & " AND    a.PostCode2 = b.PostCode2(+)"
    
    If adoSetOpen(strSql, adoSet) Then
        txtSname.Text = adoSet.Fields("Sname").Value & ""
        txtJumin1.Text = adoSet.Fields("Jumin1").Value & ""
        txtJumin2.Text = adoSet.Fields("Jumin2").Value & ""
        txtBirthDate.Text = adoSet.Fields("BirthDate").Value & ""
        txtTel.Text = adoSet.Fields("Tel").Value & ""
        txtAddr.Text = Trim(adoSet.Fields("Postname1").Value & "") & " " & _
                       Trim(adoSet.Fields("Postname2").Value & "") & " " & _
                       Trim(adoSet.Fields("Postname3").Value & "") & " " & _
                       Trim(adoSet.Fields("Juso").Value & "")
        Call adoSetClose(adoSet)
    End If
    
    txtAgeYY.Text = SetAge_Check(txtJumin1.Text, txtJumin2.Text)
    If Trim(txtSex.Text) = "" Then
        Select Case Left(txtJumin2.Text, 1)
            Case "1", "3", "0": txtSex.Text = "M"
            Case "2", "4", "9": txtSex.Text = "F"
        End Select
    End If
    Return
    
    
TextBox_Display_Clear:
    txtSname.Text = ""
    txtSex.Text = ""
    txtAgeYY.Text = ""
    txtTel.Text = ""
    txtAddr.Text = ""
    txtJumin1.Text = ""
    txtJumin2.Text = ""
    
    Return
    
End Sub
