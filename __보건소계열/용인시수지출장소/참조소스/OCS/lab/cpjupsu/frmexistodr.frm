VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#2.0#0"; "TAB32X20.OCX"
Begin VB.Form frmExistOdr 
   Caption         =   "접수환자내역 조회"
   ClientHeight    =   8445
   ClientLeft      =   75
   ClientTop       =   750
   ClientWidth     =   11745
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
   ScaleHeight     =   8445
   ScaleWidth      =   11745
   WindowState     =   2  '최대화
   Begin Threed.SSPanel SSPanel1 
      Height          =   1320
      Left            =   135
      TabIndex        =   1
      Top             =   135
      Width           =   4605
      _Version        =   65536
      _ExtentX        =   8123
      _ExtentY        =   2328
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
      Begin VB.Frame Frame1 
         Caption         =   "외.입구분"
         Height          =   600
         Left            =   225
         TabIndex        =   4
         Top             =   540
         Width           =   2535
         Begin VB.CheckBox chkOpd 
            Caption         =   "외래"
            Height          =   225
            Left            =   450
            TabIndex        =   6
            Top             =   270
            Value           =   1  '확인
            Width           =   870
         End
         Begin VB.CheckBox chkIpd 
            Caption         =   "입원"
            Height          =   225
            Left            =   1440
            TabIndex        =   5
            Top             =   270
            Width           =   870
         End
      End
      Begin MSComCtl2.DTPicker dtToDay 
         Height          =   285
         Left            =   1350
         TabIndex        =   2
         Top             =   180
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24707075
         CurrentDate     =   36413
      End
      Begin MSForms.CommandButton cmdQryAll 
         Height          =   375
         Left            =   2880
         TabIndex        =   7
         Top             =   630
         Width           =   1365
         Caption         =   "조회확인"
         Size            =   "2408;661"
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
         Left            =   315
         TabIndex        =   3
         Top             =   225
         Width           =   1005
      End
   End
   Begin FPSpreadADO.fpSpread sprOrder 
      Height          =   6315
      Left            =   90
      TabIndex        =   0
      Top             =   1575
      Width           =   11715
      _Version        =   196608
      _ExtentX        =   20664
      _ExtentY        =   11139
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
      SpreadDesigner  =   "frmExistOdr.frx":0000
      Appearance      =   1
      TextTip         =   1
      ScrollBarTrack  =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   1320
      Left            =   4950
      TabIndex        =   8
      Top             =   135
      Width           =   6720
      _Version        =   65536
      _ExtentX        =   11853
      _ExtentY        =   2328
      _StockProps     =   15
      Caption         =   "SSPanel2"
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
      Begin TabproLib.vaTabPro vaTabPro1 
         Height          =   1095
         Left            =   225
         TabIndex        =   9
         Top             =   90
         Width           =   6225
         _Version        =   131072
         _ExtentX        =   10980
         _ExtentY        =   1931
         _StockProps     =   100
         TabsPerRow      =   5
         TabCount        =   5
         Tab             =   1
         AlignTextV      =   1
         TabShape        =   3
         ApplyTo         =   2
         OffsetFromClientTop=   -1  'True
         ChamferedWidth  =   1
         ChamferedHeight =   1
         BookCornerGuardWidth=   105
         BookCornerGuardLength=   405
         TabShapeApplyTo =   1
         ThreeDInnerWidthActive=   0
         TabCaption      =   "frmExistOdr.frx":45E2
         Begin VB.ComboBox cmbRoom 
            Enabled         =   0   'False
            Height          =   300
            Left            =   -19109
            Style           =   2  '드롭다운 목록
            TabIndex        =   13
            Top             =   -15794
            Width           =   915
         End
         Begin VB.ComboBox cmbWard 
            Enabled         =   0   'False
            Height          =   300
            Left            =   -17714
            Style           =   2  '드롭다운 목록
            TabIndex        =   12
            Top             =   -15794
            Width           =   1860
         End
         Begin VB.ComboBox cmbDept 
            Enabled         =   0   'False
            Height          =   300
            Left            =   -18974
            Style           =   2  '드롭다운 목록
            TabIndex        =   11
            Top             =   -15839
            Width           =   2265
         End
         Begin VB.TextBox txtEptno 
            Height          =   330
            Left            =   2385
            TabIndex        =   10
            Top             =   540
            Width           =   1410
         End
         Begin MSComCtl2.DTPicker dtToOrder 
            Height          =   285
            Left            =   -19469
            TabIndex        =   14
            Top             =   -15824
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   24707075
            CurrentDate     =   36413
         End
         Begin MSComCtl2.DTPicker dtFrOrder 
            Height          =   285
            Left            =   -18029
            TabIndex        =   15
            Top             =   -15824
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   24707075
            CurrentDate     =   36413
         End
         Begin MSForms.CommandButton cmdQryOrderDt 
            Height          =   375
            Left            =   -20954
            TabIndex        =   24
            Top             =   -15869
            Width           =   1320
            VariousPropertyBits=   25
            Caption         =   "조회확인"
            Size            =   "2328;661"
            FontName        =   "굴림체"
            FontEffects     =   1073750016
            FontHeight      =   180
            FontCharSet     =   129
            FontPitchAndFamily=   18
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdQryWard 
            Height          =   375
            Left            =   -20954
            TabIndex        =   23
            Top             =   -15869
            Width           =   1320
            VariousPropertyBits=   25
            Caption         =   "조회확인"
            Size            =   "2328;661"
            FontName        =   "굴림체"
            FontEffects     =   1073750016
            FontHeight      =   180
            FontCharSet     =   129
            FontPitchAndFamily=   18
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdQryDept 
            Height          =   375
            Left            =   -20774
            TabIndex        =   22
            Top             =   -15914
            Width           =   1320
            VariousPropertyBits=   25
            Caption         =   "조회확인"
            Size            =   "2328;661"
            FontName        =   "굴림체"
            FontEffects     =   1073750016
            FontHeight      =   180
            FontCharSet     =   129
            FontPitchAndFamily=   18
            ParagraphAlign  =   3
         End
         Begin VB.Label Label6 
            Caption         =   "Order일자:From/To"
            Enabled         =   0   'False
            Height          =   240
            Left            =   -16679
            TabIndex        =   21
            Top             =   -15824
            Width           =   1590
         End
         Begin VB.Label Label4 
            Caption         =   "병실"
            Enabled         =   0   'False
            Height          =   195
            Left            =   -18209
            TabIndex        =   20
            Top             =   -15734
            Width           =   420
         End
         Begin VB.Label Label3 
            Caption         =   "병동"
            Enabled         =   0   'False
            Height          =   195
            Left            =   -15824
            TabIndex        =   19
            Top             =   -15734
            Width           =   465
         End
         Begin VB.Label Label2 
            Caption         =   "의뢰과선택"
            Enabled         =   0   'False
            Height          =   195
            Left            =   -16634
            TabIndex        =   18
            Top             =   -15779
            Width           =   1050
         End
         Begin VB.Label Label1 
            Caption         =   "등록번호"
            Height          =   195
            Left            =   1440
            TabIndex        =   17
            Top             =   585
            Width           =   870
         End
         Begin MSForms.CommandButton cmdExistOk 
            Height          =   375
            Left            =   4320
            TabIndex        =   16
            Top             =   495
            Width           =   1320
            Caption         =   "조회확인"
            Size            =   "2328;661"
            FontName        =   "굴림체"
            FontHeight      =   180
            FontCharSet     =   129
            FontPitchAndFamily=   18
            ParagraphAlign  =   3
         End
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmExistOdr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbWard_Click()
    Dim sWdCode     As String * 6
    
    
    If cmbWard.ListIndex = -1 Then Exit Sub
    
    sWdCode = Left(cmbWard.Text, 6)
    cmbRoom.Clear
    
    strSql = ""
    strSql = strSql & " SELECT RoomCode"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWBAS_Room"
    strSql = strSql & " WHERE  WardCode = '" & sWdCode & "'"
    strSql = strSql & " Order  By RoomCode"
    
'o  If False = adoSetOpen(strSql, adoSet) Then ReTurn
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    
    Do Until adoSet.EOF
        cmbRoom.AddItem adoSet.Fields("RoomCode").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    
End Sub

Private Sub cmdExistOk_Click()
    Dim sToDayDt            As String
    Dim sCompare            As String

    Screen.MousePointer = vbHourglass
    
    
    GoSub Form_Clear_Sub
    sToDayDt = Format(dtToDay.Value, "yyyy-MM-dd")
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
    strSql = strSql & " WHERE  a.CollDate  =   TO_DATE('" & sToDayDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.JeobsuYN  =  '*'"
    strSql = strSql & " AND    a.SLipno1  <   50"

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
    strSql = " SELECT NAME FROM TW_MIS_PMPA.TWBAS_Pass WHERE iDNumber = '" & sJcode(0) & "'"
    If False = adoSetOpen(strSql, adoPass) Then
        Return
    End If
    
    sprOrder.Col = 36:  sprOrder.Text = adoPass.Fields("Name").Value & ""
    sJcode(1) = sprOrder.Text
    
    Call adoSetClose(adoPass)
    
    
    Return
    

End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub cmdQryAll_Click()
    Dim sToDayDt            As String
    Dim sCompare            As String

    Screen.MousePointer = vbHourglass
    
    
    GoSub Form_Clear_Sub
    sToDayDt = Format(dtToDay.Value, "yyyy-MM-dd")
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
    strSql = strSql & " WHERE  a.CollDate  =   TO_DATE('" & sToDayDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.JeobsuYN  =  '*'"
    strSql = strSql & " AND    a.SLipno1  <   50"
    
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
    strSql = " SELECT NAME FROM TW_MIS_PMPA.TWBAS_Pass WHERE iDNumber = '" & sJcode(0) & "'"
    If False = adoSetOpen(strSql, adoPass) Then
        Return
    End If
    
    sprOrder.Col = 36:  sprOrder.Text = adoPass.Fields("Name").Value & ""
    sJcode(1) = sprOrder.Text
    
    Call adoSetClose(adoPass)
    
    
    Return

End Sub

Private Sub cmdQryDept_Click()
    Dim sToDayDt            As String
    Dim sCompare            As String

    Screen.MousePointer = vbHourglass
    
    
    GoSub Form_Clear_Sub
    sToDayDt = Format(dtToDay.Value, "yyyy-MM-dd")
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
    strSql = strSql & " WHERE  a.CollDate  =   TO_DATE('" & sToDayDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.JeobsuYN  =  '*'"
    strSql = strSql & " AND    a.SLipno1  <   50"

    If chkOPd.Value = "1" Then
        If chkIPd.Value = "0" Then
            strSql = strSql & " AND    a.Gbio      = 'O'"         '외래환자만
        End If
    Else
        If chkOPd.Value = "0" Then
            strSql = strSql & " AND    a.Gbio      = 'I'"         '입원환자만
        End If
    End If
        
    If cmbDept.ListIndex > -1 Then
        strSql = strSql & " AND  a.DeptCode = '" & Left(cmbDept.Text, 4) & "'"
    End If
    
    
    strSql = strSql & " AND    a.Ptno      = b.Ptno(+)"
    strSql = strSql & " AND    c.Codegu    = '12'"
    strSql = strSql & " AND    a.GeomchCd  = d.Code(+)"
    strSql = strSql & " AND    a.Drcode    = e.Drcode(+)"
    strSql = strSql & " AND    TO_NUMBER(c.Codeky)  = a.SLipno1"
    strSql = strSql & " AND    a.JeobsuDt  = f.JeobsuDt(+)"
    strSql = strSql & " AND    a.SLipno1   = f.SLipno1(+)"
    strSql = strSql & " AND    a.Orderno   = f.Orderno(+)"
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
    strSql = " SELECT  NAME  FROM  TW_MIS_PMPA.TWBAS_Pass WHERE iDNumber = '" & sJcode(0) & "'"
    If False = adoSetOpen(strSql, adoPass) Then
        Return
    End If
    
    sprOrder.Col = 36:  sprOrder.Text = adoPass.Fields("Name").Value & ""
    sJcode(1) = sprOrder.Text
    
    Call adoSetClose(adoPass)
    
    
    Return

End Sub

Private Sub cmdQryOrderDt_Click()
    Dim sToDayDt            As String
    Dim sCompare            As String
    Dim sFrOdrDt            As String
    Dim sToOdrDt            As String
    
    
    Screen.MousePointer = vbHourglass
    
    sFrOdrDt = Format(Me.dtFrOrder.Value, "yyyy-MM-dd")
    sToOdrDt = Format(Me.dtToOrder.Value, "yyyy-MM-dd")
    
    GoSub Form_Clear_Sub
    sToDayDt = Format(dtToDay.Value, "yyyy-MM-dd")
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
    strSql = strSql & " WHERE  a.CollDate  =   TO_DATE('" & sToDayDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.JeobsuYN  =  '*'"
    strSql = strSql & " AND    a.SLipno1  <   50"
    
    If chkOPd.Value = "1" Then
        If chkIPd.Value = "0" Then
            strSql = strSql & " AND    a.Gbio      = 'O'"         '외래환자만
        End If
    Else
        If chkOPd.Value = "0" Then
            strSql = strSql & " AND    a.Gbio      = 'I'"         '입원환자만
        End If
    End If
    
    strSql = strSql & " AND    a.OrderDt  >= TO_DATE('" & sFrOdrDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.OrderDt  <= TO_DATE('" & sToOdrDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.Ptno      = b.Ptno(+)"
    strSql = strSql & " AND    c.Codegu    = '12'"
    strSql = strSql & " AND    a.GeomchCd  = d.Code(+)"
    strSql = strSql & " AND    a.Drcode    = e.Drcode(+)"
    strSql = strSql & " AND    TO_NUMBER(c.Codeky)  = a.SLipno1"
    strSql = strSql & " AND    a.JeobsuDt  = f.JeobsuDt(+)"
    strSql = strSql & " AND    a.SLipno1   = f.SLipno1(+)"
    strSql = strSql & " AND    a.Orderno   = f.Orderno(+)"
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

Private Sub cmdQryWard_Click()
    Dim sToDayDt            As String
    Dim sCompare            As String

    Screen.MousePointer = vbHourglass
    
    
    GoSub Form_Clear_Sub
    sToDayDt = Format(dtToDay.Value, "yyyy-MM-dd")
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
    strSql = strSql & "        d.Codenm Samplename, e.Drname, g.Jeobsuja"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Order   a, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PATIENT  b, "
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Specode c, "
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Sample  d, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR   e, "
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_Room     f, "
    strSql = strSql & "        TWEXAM_General g  "
    strSql = strSql & " WHERE  a.CollDate  =   TO_DATE('" & sToDayDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.JeobsuYN  =  '*'"
    strSql = strSql & " AND    a.SLipno1  <   50"
    
    If chkOPd.Value = "1" Then
        If chkIPd.Value = "0" Then
            strSql = strSql & " AND    a.Gbio      = 'O'"         '외래환자만
        End If
    Else
        If chkOPd.Value = "0" Then
            strSql = strSql & " AND    a.Gbio      = 'I'"         '입원환자만
        End If
    End If
    
    If cmbWard.ListIndex > -1 Then
        If cmbRoom.ListIndex > -1 Then
            strSql = strSql & " AND  a.RoomCode  =  '" & cmbRoom.Text & "'"
            strSql = strSql & " AND  a.RoomCode  =  f.RoomCode(+)"
        Else
            strSql = strSql & " AND  a.RoomCode  =  f.RoomCode(+)"
            strSql = strSql & " AND  f.WardCode  =  '" & Left(cmbWard.Text, 6) & "'"
        End If
    End If
    
    strSql = strSql & " AND    a.Ptno      = b.Ptno(+)"
    strSql = strSql & " AND    c.Codegu    = '12'"
    strSql = strSql & " AND    a.GeomchCd  = d.Code(+)"
    strSql = strSql & " AND    a.Drcode    = e.Drcode(+)"
    strSql = strSql & " AND    TO_NUMBER(c.Codeky)  = a.SLipno1"
    strSql = strSql & " AND    a.JeobsuDt  = g.JeobsuDt(+)"
    strSql = strSql & " AND    a.SLipno1   = g.SLipno1(+)"
    strSql = strSql & " AND    a.Orderno   = g.Orderno(+)"
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
    strSql = " SELECT NAME FROM TW_MIS_PMPA.TWBAS_Pass WHERE iDNumber = '" & sJcode(0) & "'"
    If False = adoSetOpen(strSql, adoPass) Then
        Return
    End If
    
    sprOrder.Col = 36:  sprOrder.Text = adoPass.Fields("Name").Value & ""
    sJcode(1) = sprOrder.Text
    
    Call adoSetClose(adoPass)
    
    
    Return

End Sub

Private Sub Form_Load()
    
    If GstrIOGubun = "IPD" Then
        chkIPd.Value = "1"
        chkOPd.Value = "0"
    Else
        chkOPd.Value = "0"
        chkIPd.Value = "1"
    End If
    
    
    dtToDay.Value = Dual_Date_Get("yyyy-MM-dd")
    dtFrOrder.Value = Format(dtToDay.Value, "yyyy-MM-dd")
    dtToOrder.Value = Format(dtToDay.Value, "yyyy-MM-dd")
        
    GoSub Get_DeptData
    GoSub Get_WardData
    
    Exit Sub
    
    
Get_DeptData:
    Dim sDeptCode       As String * 4
    
    strSql = ""
    strSql = strSql & " SELECT DeptCode, DeptNamek"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWBAS_DEPT"
    strSql = strSql & " Order  By PRintRanking"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sDeptCode = adoSet.Fields("DeptCode").Value & ""
        cmbDept.AddItem sDeptCode & "" & Trim(adoSet.Fields("DeptnameK").Value & "")
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
Get_WardData:
    Dim sWardCode       As String * 6
    
    strSql = ""
    strSql = strSql & " SELECT WardCode, WardName"
    strSql = strSql & " FROM   TW_MIS_PMPA.TWBAS_Ward"
    strSql = strSql & " Order  By WardCode"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sWardCode = adoSet.Fields("WardCode").Value & ""
        cmbWard.AddItem sWardCode & " " & adoSet.Fields("WardName").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
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
        txtEptno.Text = Format(txtEptno.Text, "00000000")
        
    End If
    
End Sub

