VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0"; "FPSPR30.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmQuery 
   Caption         =   "접수환자조회"
   ClientHeight    =   8010
   ClientLeft      =   330
   ClientTop       =   1185
   ClientWidth     =   11100
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
   ScaleHeight     =   8010
   ScaleWidth      =   11100
   WindowState     =   2  '최대화
   Begin VB.TextBox txtWard 
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   8100
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "txtWard"
      Top             =   1260
      Width           =   870
   End
   Begin VB.TextBox txtSname 
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   6210
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "txtSname"
      Top             =   900
      Width           =   915
   End
   Begin VB.TextBox txtSexAge 
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   7110
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "txtSexAge"
      Top             =   900
      Width           =   555
   End
   Begin VB.TextBox txtJumin 
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   7695
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Text            =   "6708151462411"
      Top             =   900
      Width           =   1320
   End
   Begin VB.TextBox txtTel 
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   9315
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "txtTel"
      Top             =   1080
      Width           =   1500
   End
   Begin VB.TextBox txtAddr 
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   6210
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Text            =   "txtAddr"
      Top             =   1575
      Width           =   4470
   End
   Begin VB.TextBox txtDept 
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   6210
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   "txtDept"
      Top             =   1260
      Width           =   870
   End
   Begin VB.TextBox txtDr 
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   7110
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   "txtDr"
      Top             =   1260
      Width           =   1005
   End
   Begin VB.TextBox txtQryPtno 
      Height          =   285
      Left            =   6210
      MaxLength       =   12
      TabIndex        =   18
      Top             =   585
      Width           =   2805
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   420
      Left            =   135
      TabIndex        =   17
      Top             =   90
      Width           =   4830
      _Version        =   65536
      _ExtentX        =   8520
      _ExtentY        =   741
      _StockProps     =   15
      Caption         =   "접수환자조회"
      ForeColor       =   65535
      BackColor       =   8388608
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
   Begin FPSpreadADO.fpSpread sprOrder 
      Height          =   5370
      Left            =   5040
      TabIndex        =   1
      Top             =   2040
      Width           =   6810
      _Version        =   196608
      _ExtentX        =   12012
      _ExtentY        =   9472
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
      MaxCols         =   8
      ScrollBars      =   2
      SpreadDesigner  =   "frmQuery.frx":0000
      Appearance      =   1
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   915
      Left            =   135
      TabIndex        =   0
      Top             =   585
      Width           =   4830
      _Version        =   65536
      _ExtentX        =   8520
      _ExtentY        =   1614
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
      Begin MSComCtl2.DTPicker dtToDate 
         Height          =   330
         Left            =   2565
         TabIndex        =   8
         Top             =   135
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24510467
         CurrentDate     =   36431
      End
      Begin MSComCtl2.DTPicker dtFrDate 
         Height          =   330
         Left            =   1215
         TabIndex        =   7
         Top             =   135
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24510467
         CurrentDate     =   36431
      End
      Begin Threed.SSPanel panelVerify 
         Height          =   285
         Left            =   2790
         TabIndex        =   4
         Top             =   540
         Visible         =   0   'False
         Width           =   1905
         _Version        =   65536
         _ExtentX        =   3360
         _ExtentY        =   503
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
         Begin VB.OptionButton optNoVer 
            Caption         =   "미확인"
            Height          =   180
            Left            =   855
            TabIndex        =   6
            Top             =   45
            Width           =   870
         End
         Begin VB.OptionButton optVer 
            Caption         =   "확인"
            Height          =   180
            Left            =   135
            TabIndex        =   5
            Top             =   45
            Value           =   -1  'True
            Width           =   690
         End
      End
      Begin VB.OptionButton optIpd 
         Caption         =   "입 원"
         Height          =   180
         Left            =   2025
         TabIndex        =   3
         Top             =   585
         Width           =   780
      End
      Begin VB.OptionButton optOpd 
         Caption         =   "외 래"
         Height          =   180
         Left            =   1170
         TabIndex        =   2
         Top             =   585
         Value           =   -1  'True
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "기준일자"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   180
         Width           =   735
      End
   End
   Begin Threed.SSPanel panelOPd 
      Height          =   5370
      Left            =   90
      TabIndex        =   12
      Top             =   2070
      Width           =   4920
      _Version        =   65536
      _ExtentX        =   8678
      _ExtentY        =   9472
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
      Begin FPSpreadADO.fpSpread sproOrder 
         Height          =   5145
         Left            =   45
         TabIndex        =   13
         Top             =   90
         Width           =   4830
         _Version        =   196608
         _ExtentX        =   8520
         _ExtentY        =   9075
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
         ScrollBars      =   2
         SpreadDesigner  =   "frmQuery.frx":3C19
         Appearance      =   1
      End
   End
   Begin Threed.SSPanel paneliPd 
      Height          =   5370
      Left            =   90
      TabIndex        =   10
      Top             =   2070
      Width           =   4920
      _Version        =   65536
      _ExtentX        =   8678
      _ExtentY        =   9472
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
      Begin FPSpreadADO.fpSpread spriOrder 
         Height          =   5145
         Left            =   45
         TabIndex        =   11
         Top             =   90
         Width           =   4830
         _Version        =   196608
         _ExtentX        =   8520
         _ExtentY        =   9075
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
         MaxCols         =   8
         ScrollBars      =   2
         SpreadDesigner  =   "frmQuery.frx":79A6
         Appearance      =   1
      End
   End
   Begin VB.Label Label2 
      Caption         =   "등록번호"
      Height          =   240
      Left            =   5400
      TabIndex        =   31
      Top             =   630
      Width           =   780
   End
   Begin VB.Label Label7 
      Caption         =   "☎"
      Height          =   195
      Left            =   9045
      TabIndex        =   30
      Top             =   1125
      Width           =   240
   End
   Begin VB.Label Label8 
      Caption         =   "주소"
      Height          =   195
      Left            =   5445
      TabIndex        =   29
      Top             =   1620
      Width           =   600
   End
   Begin VB.Label Label9 
      Caption         =   "과/의사"
      Height          =   195
      Left            =   5445
      TabIndex        =   28
      Top             =   1305
      Width           =   690
   End
   Begin MSForms.CommandButton cmdCls 
      Height          =   465
      Left            =   9315
      TabIndex        =   27
      Top             =   585
      Width           =   1500
      Caption         =   "Clear"
      PicturePosition =   327683
      Size            =   "2646;820"
      Picture         =   "frmQuery.frx":B661
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   465
      Left            =   3285
      TabIndex        =   16
      Top             =   1575
      Width           =   1680
      Caption         =   "Clear"
      PicturePosition =   327683
      Size            =   "2963;820"
      Picture         =   "frmQuery.frx":CDF3
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdoOrder 
      Height          =   465
      Left            =   1575
      TabIndex        =   14
      Top             =   1575
      Width           =   1680
      Caption         =   "조회확인(o)"
      PicturePosition =   327683
      Size            =   "2963;820"
      Picture         =   "frmQuery.frx":E585
      FontName        =   "굴림체"
      FontHeight      =   180
      FontCharSet     =   129
      FontPitchAndFamily=   18
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdiOrder 
      Height          =   465
      Left            =   1575
      TabIndex        =   15
      Top             =   1575
      Visible         =   0   'False
      Width           =   1680
      Caption         =   "조회확인(i)"
      PicturePosition =   327683
      Size            =   "2963;820"
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
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sFlagPtnoBoxEnter        As String

Public Function convLabnoToPtno(ByVal sLabno As String) As String
    Dim sJeobsuDt       As String
    Dim iSLipno1        As Integer
    Dim iSLipno2        As Integer
    Dim adoFind         As ADODB.Recordset
    
    
    convLabnoToPtno = ""
    
    'Labno 를 Ptno 로 바꾸어 Query (임상병리과 김일형)
    
    sJeobsuDt = convLabnoToExpand(Left(sLabno, 5))
    iSLipno1 = Val(Mid(sLabno, 6, 2))
    iSLipno2 = Val(Mid(sLabno, 8, 5))
    
    strSql = ""
    strSql = strSql & " SELECT Ptno"
    strSql = strSql & " FROM   TWEXAM_General"
    strSql = strSql & " WHERE  JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND    SLipno1  = " & iSLipno1
    strSql = strSql & " AND    SLipno2  = " & iSLipno2
    If False = adoSetOpen(strSql, adoFind) Then
        Exit Function
    End If
    
    
    convLabnoToPtno = adoFind.Fields("Ptno").Value & ""
    
    Call adoSetClose(adoFind)
    
    
End Function




Private Sub CmdCancel_Click()
    
    Call Spread_Set_Clear(sproOrder)
    Call Spread_Set_Clear(spriOrder)
    Call Spread_Set_Clear(sprOrder)
    Call cmdCls_Click
    
End Sub

Private Sub cmdCls_Click()
    txtQryPtno.Text = ""
    txtWard.Text = ""
    txtSname.Text = ""
    txtSexage.Text = ""
    txtJumin.Text = ""
    txtTel.Text = ""
    txtDept.Text = ""
    txtDr.Text = ""
    txtAddr.Text = ""

End Sub

Public Sub cmdiOrder_Click()
    Dim sGbio       As String
    Dim sVerify     As String
    Dim sFrDate     As String
    Dim sToDate     As String
    
    
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    Call Spread_Set_Clear(spriOrder)
    GoSub Set_WhereData
    GoSub Get_Main_Proc
    If spriOrder.DataRowCnt > 0 Then
        Call spriOrder_DblClick(3, 1)
    End If
    
    Exit Sub
    
    

Set_WhereData:
    If optOpd.Value = True Then
        sGbio = "O"
        sVerify = ""
    End If
    
    If optIpd.Value = True Then
        sGbio = "I"
        If optVer.Value = True Then sVerify = "Y"
        If optNoVer.Value = True Then sVerify = "12"
    End If
    
    Return
    
Get_Main_Proc:
    'strSql = ""
    'strSql = strSql & "  SELECT /*+ IDNEX (TW_MIS_PMPA.TWBAS_PATIENT INDEX_PATIENT0) */        "
    
    strSql = ""
    strSql = strSql & " SELECT  a.RoomCode, "
    strSql = strSql & "         TO_CHAR(a.JeobsuDt, 'yyyy-MM-dd') JeobsuDt,"
    strSql = strSql & "         TO_CHAR(a.OrderDt,  'yyyy-MM-dd') OrderDt, "
    strSql = strSql & "         a.JeobsuT1, a.JeobsuT2,"
    strSql = strSql & "         a.PTNO, b.Sname"
    strSql = strSql & "  FROM   TWEXAM_GENERAL a,"
    strSql = strSql & "         TW_MIS_PMPA.TWBAS_PATIENT  b"
    strSql = strSql & " WHERE   a.JeobsuDt >= TO_DATE('" & sFrDate & "','YYYY-MM-DD')"
    strSql = strSql & " AND     a.JeobsuDt <= TO_DATE('" & sToDate & "','YYYY-MM-DD')"
    strSql = strSql & " AND     a.GbIO      = '" & sGbio & "'"
'C    strSql = strSql & " AND     a.SLipno1  <  52    "
    strSql = strSql & " AND     a.SLipno1  <  90    "
    If sFlagPtnoBoxEnter = "13" Then
        If Trim(txtQryPtno.Text) <> "" Then
            strSql = strSql & " AND  a.Ptno    = '" & txtQryPtno.Text & "'"
        End If
    End If

    If sVerify = "Y" Then
        strSql = strSql & " AND  a.GbCH    = 'Y'"
    Else
        strSql = strSql & " AND  a.GbCh   IN ('1','2')"
    End If
    
    strSql = strSql & "  AND    a.Ptno = b.Ptno(+)"
    strSql = strSql & "  GROUP BY  a.JeobsuDt, a.JeobsuT1, a.JeobsuT2, a.OrderDt, a.PTNO, b.Sname,a.ROOMCODE"
    strSql = strSql & "  ORDER BY  a.JeobsuDt, a.JeobsuT1, a.JeobsuT2"

    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        spriOrder.Row = spriOrder.DataRowCnt + 1
        
        spriOrder.Col = 2: spriOrder.Text = adoSet.Fields("RoomCode").Value & ""
        spriOrder.Col = 3: spriOrder.Text = adoSet.Fields("JeobsuDt").Value & ""
        spriOrder.Col = 4: spriOrder.Text = adoSet.Fields("JeobsuT1").Value & ""
        spriOrder.Col = 5: spriOrder.Text = adoSet.Fields("JeobsuT2").Value & ""
        spriOrder.Col = 6: spriOrder.Text = adoSet.Fields("OrderDt").Value & " " & _
                                            Format(adoSet.Fields("JeobsuT1").Value, "00") & ":" & _
                                            Format(adoSet.Fields("JeobsuT2").Value, "00") & ""
        spriOrder.Col = 7: spriOrder.Text = adoSet.Fields("Ptno").Value & ""
        spriOrder.Col = 8: spriOrder.Text = adoSet.Fields("Sname").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return
    
    
End Sub

Public Sub cmdoOrder_Click()
    Dim sGbio       As String
    Dim sVerify     As String
    Dim sFrDate     As String
    Dim sToDate     As String
    
    
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
    
    Call Spread_Set_Clear(sproOrder)
    GoSub Set_WhereData
    GoSub Get_Main_Proc
    If sproOrder.DataRowCnt > 0 Then
        Call sproOrder_DblClick(3, 1)
    End If
    Exit Sub
    
    

Set_WhereData:
    If optOpd.Value = True Then
        sGbio = "O"
        sVerify = "Y"
    End If
    
    If optIpd.Value = True Then
        sGbio = "I"
        If optVer.Value = True Then sVerify = "Y"
        If optNoVer.Value = True Then sVerify = "12"
    End If
    
    Return
    
    
    
Get_Main_Proc:
    
    'strSql = strSql & "  SELECT /*+ IDNEX (TW_MIS_PMPA.TWBAS_PATIENT INDEX_PATIENT0) */ "
    
    
    strSql = ""
    strSql = strSql & " SELECT  a.DeptCode, c.DeptNamek,"
    strSql = strSql & "         TO_CHAR(a.JeobsuDt, 'YYYY-MM-DD') JeobsuDt,"
    strSql = strSql & "         TO_CHAR(a.OrderDt,  'yyyy-MM-dd') OrderDt, "
    strSql = strSql & "         a.JeobsuT1, a.JeobsuT2,"
    strSql = strSql & "         a.PTNO, b.Sname"
    strSql = strSql & "  FROM   TW_MIS_EXAM.TWEXAM_General a,"
    strSql = strSql & "         TW_MIS_PMPA.TWBAS_PATIENT  b,"
    strSql = strSql & "         TW_MIS_PMPA.TWBAS_DEPT     c"
    strSql = strSql & " WHERE   a.JeobsuDt >= TO_DATE('" & sFrDate & "','YYYY-MM-DD')"
    strSql = strSql & " AND     a.JeobsuDt <= TO_DATE('" & sToDate & "','YYYY-MM-DD')"
'C    strSql = strSql & " AND     a.SLipno1  < 52"
    strSql = strSql & " AND     a.SLipno1  < 90"
    strSql = strSql & " AND     a.GbIO      = '" & sGbio & "'"
    
    If sFlagPtnoBoxEnter = "13" Then
        If Trim(txtQryPtno.Text) <> "" Then
            strSql = strSql & " AND  a.Ptno    = '" & txtQryPtno.Text & "'"
        End If
    End If

    If sVerify = "Y" Then
        strSql = strSql & " AND  a.GbCH    = 'Y'"
    Else
        strSql = strSql & " AND  a.GbCh   IN ('1','2')"
    End If
    
    strSql = strSql & "  AND    a.Ptno     = b.Ptno(+)"
    strSql = strSql & "  AND    a.DeptCode = c.DeptCode(+)"
    strSql = strSql & "  GROUP BY  a.JeobsuDt, a.JeobsuT1, a.JeobsuT2, a.OrderDt, a.PTNO, b.Sname,a.DeptCode, c.Deptnamek"
    strSql = strSql & "  ORDER BY  a.JeobsuDt, a.JeobsuT1, a.JeobsuT2"

    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sproOrder.Row = sproOrder.DataRowCnt + 1
        
        sproOrder.Col = 2: sproOrder.Text = adoSet.Fields("DeptCode").Value & ""
        sproOrder.Col = 3: sproOrder.Text = adoSet.Fields("DeptnameK").Value & ""
        sproOrder.Col = 4: sproOrder.Text = adoSet.Fields("JeobsuDt").Value & ""
        sproOrder.Col = 5: sproOrder.Text = adoSet.Fields("JeobsuT1").Value & ""
        sproOrder.Col = 6: sproOrder.Text = adoSet.Fields("JeobsuT2").Value & ""
        sproOrder.Col = 7: sproOrder.Text = adoSet.Fields("OrderDt").Value & ""
        sproOrder.Col = 8: sproOrder.Text = adoSet.Fields("Ptno").Value & ""
        sproOrder.Col = 9: sproOrder.Text = adoSet.Fields("Sname").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return


Get_Main_Proc1:
    'strSql = ""
    'strSql = strSql & "  SELECT /*+ IDNEX (TW_MIS_PMPA.TWBAS_PATIENT INDEX_PATIENT0) */ "
    
    strSql = ""
    strSql = strSql & " SELECT  a.DeptCode, c.DeptNamek,"
    strSql = strSql & "         TO_CHAR(a.COLLDate, 'YYYY-MM-DD') COLLDate,"
    strSql = strSql & "         TO_CHAR(a.OrderDt,  'yyyy-MM-dd') OrderDt, "
    strSql = strSql & "         a.COLLHH, a.COLLMM,"
    strSql = strSql & "         a.PTNO, b.Sname"
    strSql = strSql & "  FROM   TW_MIS_EXAM.TWEXAM_Order a,"
    strSql = strSql & "         TW_MIS_PMPA.TWBAS_PATIENT  b,"
    strSql = strSql & "         TW_MIS_PMPA.TWBAS_DEPT     c"
    strSql = strSql & " WHERE   a.COLLDate >= TO_DATE('" & sFrDate & "','YYYY-MM-DD')"
    strSql = strSql & " AND     a.COLLDate <= TO_DATE('" & sToDate & "','YYYY-MM-DD')"
'C    strSql = strSql & " AND     a.SLipno1  < 52"
    strSql = strSql & " AND     a.SLipno1  < 90 "
    strSql = strSql & " AND     a.GbIO      = '" & sGbio & "'"
    
    If sFlagPtnoBoxEnter = "13" Then
        If Trim(txtQryPtno.Text) <> "" Then
            strSql = strSql & " AND  a.Ptno    = '" & txtQryPtno.Text & "'"
        End If
    End If

    If sVerify = "Y" Then
        strSql = strSql & " AND  a.GbCH    = 'Y'"
    Else
        strSql = strSql & " AND  a.GbCh   IN ('1','2')"
    End If
    
    strSql = strSql & "  AND    a.Ptno     = b.Ptno(+)"
    strSql = strSql & "  AND    a.DeptCode = c.DeptCode(+)"
    strSql = strSql & "  GROUP BY  a.COLLDate, a.COLLHH, a.COLLMM, a.OrderDt, a.PTNO, b.Sname,a.DeptCode, c.Deptnamek"
    strSql = strSql & "  ORDER BY  a.COLLDate, a.COLLHH, a.COLLMM"

    If False = adoSetOpen(strSql, adoSet) Then Return
    
    Do Until adoSet.EOF
        sproOrder.Row = sproOrder.DataRowCnt + 1
        
        sproOrder.Col = 2: sproOrder.Text = adoSet.Fields("DeptCode").Value & ""
        sproOrder.Col = 3: sproOrder.Text = adoSet.Fields("DeptnameK").Value & ""
        sproOrder.Col = 4: sproOrder.Text = adoSet.Fields("COLLDate").Value & ""
        sproOrder.Col = 5: sproOrder.Text = adoSet.Fields("COLLHH").Value & ""
        sproOrder.Col = 6: sproOrder.Text = adoSet.Fields("COLLMM").Value & ""
        sproOrder.Col = 7: sproOrder.Text = adoSet.Fields("OrderDt").Value & ""
        sproOrder.Col = 8: sproOrder.Text = adoSet.Fields("Ptno").Value & ""
        sproOrder.Col = 9: sproOrder.Text = adoSet.Fields("Sname").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    Return
End Sub


Private Sub cmdQryPtno_Click()

End Sub

Private Sub Form_Activate()
    
    Me.WindowState = vbMaximized
    sFlagPtnoBoxEnter = ""
    
End Sub

Private Sub Form_Load()
    
    
    Me.dtFrDate.Value = Dual_Date_Get("yyyy-MM-dd")
    Me.dtToDate.Value = Dual_Date_Get("yyyy-MM-dd")
    
    Me.txtQryPtno.Text = ""
    Me.txtWard.Text = ""
    Me.txtSname.Text = ""
    Me.txtSexage.Text = ""
    Me.txtJumin.Text = ""
    Me.txtTel.Text = ""
    Me.txtDept.Text = ""
    Me.txtDr.Text = ""
    Me.txtAddr.Text = ""
    
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

Private Sub optIpd_Click()
    
    If optIpd.Value = True Then
        Call CmdCancel_Click
        panelVerify.Visible = True
        panelOpd.Visible = False
        paneliPd.Visible = True
        paneliPd.ZOrder 0
        
        cmdoOrder.Visible = False
        cmdiOrder.Visible = True
        cmdiOrder.ZOrder 0
        
    End If
        
        
    
End Sub

Private Sub optOpd_Click()
    
    If optOpd.Value = True Then
        Call CmdCancel_Click
        panelVerify.Visible = False
        paneliPd.Visible = False
        panelOpd.Visible = True
        panelOpd.ZOrder 0
        
        cmdiOrder.Visible = False
        cmdoOrder.Visible = True
        cmdoOrder.ZOrder 0
        
    End If
    

End Sub

Private Sub spriOrder_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
    If Col = 1 Then
        spriOrder.Row = Row
        spriOrder.Col = 3: GLabelJeobsuDt = spriOrder.Text
        spriOrder.Col = 7: GLabelPtno = spriOrder.Text
        spriOrder.Col = 3: GLabelJDt = spriOrder.Text
        spriOrder.Col = 4: GLabelJT1 = spriOrder.Text
        spriOrder.Col = 5: GLabelJT2 = spriOrder.Text
        GLabelLoadCheck = "LOAD"
        frmBarCode.Show vbModal
    End If
    
End Sub

Private Sub spriOrder_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim sQryPtno        As String
    Dim sQryJeobsuDt    As String
    Dim sQryJeobsuT1    As String
    Dim sQryJeobsuT2    As String
    
    Dim sGbio       As String
    Dim sVerify     As String
    Dim sFrDate     As String
    Dim sToDate     As String
    
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
        
        
        
    If Row = 0 Then Exit Sub
    If Row > spriOrder.DataRowCnt Then Exit Sub
    
    spriOrder.Row = Row
    spriOrder.Col = 7: sQryPtno = spriOrder.Text
    spriOrder.Col = 3: sQryJeobsuDt = spriOrder.Text
    spriOrder.Col = 4: sQryJeobsuT1 = spriOrder.Text
    spriOrder.Col = 5: sQryJeobsuT2 = spriOrder.Text
    
    Call Spread_Set_Clear(sprOrder)
    GoSub Set_WhereData
    GoSub Get_General_iOrder
    GoSub Get_Patient_Data
    
    Exit Sub
    
    
Set_WhereData:
    If optOpd.Value = True Then
        sGbio = "O"
        sVerify = "Y"
    End If
    
    If optIpd.Value = True Then
        sGbio = "I"
        If optVer.Value = True Then sVerify = "Y"
        If optNoVer.Value = True Then sVerify = "12"
    End If
    
    Return
    
Get_General_iOrder:
    strSql = ""
    strSql = strSql & "  SELECT a.ItemCd Codeky, c.ItemNM ItemName, a.GeomchCd, d.Codenm Samplename, a.Orderno,"
    strSql = strSql & "         TO_CHAR(b.GeomsaDt, 'YYYY-MM-DD') GEOMSADT,"
    strSql = strSql & "         b.GeomsaT1, b.GeomsaT2, e.Name,"
    strSql = strSql & "         TO_CHAR(f.GBDate,'yyyy-MM-dd hh24:mi') GBDate,"
    strSql = strSql & "         TO_CHAR(a.JeobsuDt, 'yyyy-MM-dd') JeobsuDt"
    strSql = strSql & "  FROM   TWEXAM_General_Sub a,        "
    strSql = strSql & "         TWEXAM_General     b,        "
    strSql = strSql & "         TWEXAM_itemML      c,        "
    strSql = strSql & "         TWEXAM_Sample      d,        "
    strSql = strSql & "         TW_MIS_PMPA.TWBas_Pass e,    "
    strSql = strSql & "         TWEXAM_Order       f         "
    strSql = strSql & "  WHERE  a.Ptno     =  '" & sQryPtno & "'"
    strSql = strSql & "  AND    a.JeobsuDt  = TO_DATE('" & sQryJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & "  AND    b.JeobsuT1  = '" & sQryJeobsuT1 & "'"
    strSql = strSql & "  AND    b.JeobsuT2  = '" & sQryJeobsuT2 & "'"
    strSql = strSql & "  AND    b.GBIO     = '" & sGbio & "'"
    strSql = strSql & "  AND    a.Routincd  = a.itemCd "
    strSql = strSql & "  AND    a.JeobsuDt  = b.JeobsuDt(+) "
    strSql = strSql & "  AND    a.SLipno1   = b.SLipno1(+) "
    strSql = strSql & "  AND    a.SLipno2   = b.SLipno2(+) "
    strSql = strSql & "  AND    a.JeobsuDt  = f.CollDate(+)"
    strSql = strSql & "  AND    a.Orderno   = f.Orderno(+)"
    strSql = strSql & "  AND    a.ItemCd    = c.Codeky"
    strSql = strSql & "  AND    a.GeomchCd  = d.Code(+)"
    
    If sVerify = "Y" Then
        strSql = strSql & " AND  f.GbCH    = 'Y'"
    Else
        strSql = strSql & " AND  f.GbCh   IN ('1','2')"
    End If
    
    strSql = strSql & "  AND    f.CollID   = e.Idnumber(+) "
    strSql = strSql & "  AND   ( e.ProgramID IS NULL Or e.ProgramID = ' ')"
    strSql = strSql & "  UNION ALL"
    strSql = strSql & "  SELECT distinct a.Routincd Codeky, c.Routinnm ItemName, a.GeomchCd, d.Codenm Samplename, a.Orderno,"
    strSql = strSql & "         TO_CHAR(b.GeomsaDt, 'YYYY-MM-DD') GEOMSADT,"
    strSql = strSql & "         b.GeomsaT1, b.GeomsaT2, e.Name,"
    strSql = strSql & "         TO_CHAR(f.GBDate,'yyyy-MM-dd hh24:mi') GBDate,"
    strSql = strSql & "         TO_CHAR(a.JeobsuDt, 'yyyy-MM-dd') JeobsuDt"
    strSql = strSql & "  FROM   TWEXAM_General_Sub a,        "
    strSql = strSql & "         TWEXAM_General     b,        "
    strSql = strSql & "         TW_MIS_EXAM.TWEXAM_Routine     c,        "
    strSql = strSql & "         TW_MIS_EXAM.TWEXAM_Sample      d,        "
    strSql = strSql & "         TW_MIS_PMPA.TWBas_Pass         e,        "
    strSql = strSql & "         TW_MIS_EXAM.TWEXAM_Order       f         "
    strSql = strSql & "  WHERE  a.Ptno     =  '" & sQryPtno & "'"
    strSql = strSql & "  AND    a.JeobsuDt  = TO_DATE('" & sQryJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & "  AND    b.JeobsuT1  = '" & sQryJeobsuT1 & "'"
    strSql = strSql & "  AND    b.JeobsuT2  = '" & sQryJeobsuT2 & "'"
    strSql = strSql & "  AND    b.GBIO     = '" & sGbio & "'"
    strSql = strSql & "  AND    a.Routincd != a.itemCd "
    strSql = strSql & "  AND    a.JeobsuDt  = b.JeobsuDt(+) "
    strSql = strSql & "  AND    a.SLipno1   = b.SLipno1(+) "
    strSql = strSql & "  AND    a.SLipno2   = b.SLipno2(+) "
    strSql = strSql & "  AND    a.JeobsuDt  = f.CollDate(+)"
    strSql = strSql & "  AND    a.Orderno   = f.Orderno(+)"
    strSql = strSql & "  AND    a.RoutinCD  = c.RoutinCd"
    strSql = strSql & "  AND    a.GeomchCd  = d.Code(+)"
    If sVerify = "Y" Then
        strSql = strSql & " AND  f.GbCH    = 'Y'"
    Else
        strSql = strSql & " AND  f.GbCh   IN ('1','2')"
    End If
    strSql = strSql & "  AND    f.CollID   = e.Idnumber(+) "
    strSql = strSql & "  AND   ( e.ProgramID IS NULL Or e.ProgramID = ' ')"
    strSql = strSql & "  order by Codeky"
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    Do Until adoSet.EOF
        sprOrder.Row = sprOrder.DataRowCnt + 1
        sprOrder.Col = 1: sprOrder.Text = Left(adoSet.Fields("Codeky").Value & "", 2)
        sprOrder.Col = 2: sprOrder.Text = adoSet.Fields("Codeky").Value & ""
        sprOrder.Col = 3: sprOrder.Text = adoSet.Fields("ItemName").Value & ""
        sprOrder.Col = 4: sprOrder.Text = adoSet.Fields("Samplename").Value & ""
        If sVerify = "Y" Then
            sprOrder.Col = 5: sprOrder.Text = adoSet.Fields("Name").Value & ""
            sprOrder.Col = 6: sprOrder.Text = adoSet.Fields("GBDate").Value & ""
            'sprOrder.Col = 6: sprOrder.Text = adoSet.Fields("GeomsaDt").Value & " " & _
            '                                  Format(adoSet.Fields("GeomsaT1").Value & "", "00") & ":" & _
            '                                  Format(adoSet.Fields("GeomsaT2").Value & "", "00")
            sprOrder.Col = 7: sprOrder.Text = adoSet.Fields("Orderno").Value & ""
        End If
        sprOrder.Col = 8: sprOrder.Text = adoSet.Fields("JeobsuDt").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
    
Get_General_Order:
    strSql = ""
    strSql = strSql & "   SELECT a.GeomChCd, c.Codenm SampleName, "
    strSql = strSql & "          a.ItemCd,   b.ItemNM ItemName,"
    strSql = strSql & "          TO_CHAR(a.CollDate,'YYYY-MM-DD') CollDate,"
    strSql = strSql & "          a.CollHH, a.CollMM, a.Collid, d.Name"
    strSql = strSql & "   FROM   TW_MIS_EXAM.TWEXAM_Order   a,"
    strSql = strSql & "          TW_MIS_EXAM.TWEXAM_itemML  b,"
    strSql = strSql & "          TW_MIS_EXAM.TWEXAM_Sample  c,"
    strSql = strSql & "          TW_MIS_PMPA.TWBas_Pass     d "
    strSql = strSql & "   WHERE  a.Ptno     =  '" & sQryPtno & "'"
    strSql = strSql & "   AND    a.CollDate = TO_DATE('" & sQryJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & "   AND    a.CollHH = '" & sQryJeobsuT1 & "'"
    strSql = strSql & "   AND    a.CollMM = '" & sQryJeobsuT2 & "'"
    strSql = strSql & "   AND    a.GBIO     = '" & sGbio & "'"
    If sVerify = "Y" Then
        strSql = strSql & " AND  a.GbCH    = 'Y'"
    Else
        strSql = strSql & " AND  a.GbCh   IN ('1','2')"
    End If
    strSql = strSql & "   AND    a.ItemCd   = b.Codeky"
    strSql = strSql & "   AND    a.GeomChCd = c.Code(+)"
    strSql = strSql & "   AND    a.Collid   = d.Idnumber(+)"
    strSql = strSql & "   AND   ( d.ProgramID IS NULL Or d.ProgramID = ' ')"
    strSql = strSql & " Union ALL"
    strSql = strSql & "   SELECT DISTINCT a.GeomChCd,  c.Codenm SampleName, "
    strSql = strSql & "                   a.ItemCd,    b.RoutinNM ItemName,"
    strSql = strSql & "          TO_CHAR(a.CollDate,'YYYY-MM-DD') CollDate,"
    strSql = strSql & "          a.CollHH, a.CollMM, a.Collid, d.Name"
    strSql = strSql & "   FROM   TW_MIS_EXAM.TWEXAM_Order   a,"
    strSql = strSql & "          TW_MIS_EXAM.TWEXAM_Routine b,"
    strSql = strSql & "          TW_MIS_EXAM.TWEXAM_Sample  c,"
    strSql = strSql & "          TW_MIS_PMPA.TWBas_Pass     d "
    strSql = strSql & "   WHERE  a.Ptno     =  '" & sQryPtno & "'"
    strSql = strSql & "   AND    a.COLLDate = TO_DATE('" & sQryJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & "   AND    a.COLLHH   = '" & sQryJeobsuT1 & "'"
    strSql = strSql & "   AND    a.COLLMM   = '" & sQryJeobsuT2 & "'"
    strSql = strSql & "   AND    a.GBIO     = '" & sGbio & "'"
    If sVerify = "Y" Then
        strSql = strSql & " AND  a.GbCH    = 'Y'"
    Else
        strSql = strSql & " AND  a.GbCh   IN ('1','2')"
    End If
    strSql = strSql & "   AND    a.ItemCd   = b.RoutinCd"
    strSql = strSql & "   AND    a.GeomChCd = c.Code(+)"
    strSql = strSql & "   AND    a.Collid   = d.Idnumber(+)"
    strSql = strSql & "   AND   ( d.ProgramID IS NULL Or d.ProgramID = ' ')"
    strSql = strSql & "   Order  By CollDate, CollHH, CollMM, iTemCd"
    
    If False = adoSetOpen(strSql, adoSet) Then Exit Sub
    Do Until adoSet.EOF
        sprOrder.Row = sprOrder.DataRowCnt + 1
        sprOrder.Col = 1: sprOrder.Text = Left(adoSet.Fields("ItemCd").Value & "", 2)
        sprOrder.Col = 2: sprOrder.Text = adoSet.Fields("ItemCD").Value & ""
        sprOrder.Col = 3: sprOrder.Text = adoSet.Fields("ItemName").Value & ""
        sprOrder.Col = 4: sprOrder.Text = adoSet.Fields("Samplename").Value & ""
        If sVerify = "Y" Then
            sprOrder.Col = 5: sprOrder.Text = adoSet.Fields("Name").Value & ""
            sprOrder.Col = 6: sprOrder.Text = adoSet.Fields("CollDate").Value & " " & _
                                              Format(adoSet.Fields("CollHH").Value & "", "00") & ":" & _
                                              Format(adoSet.Fields("CollMM").Value & "", "00")
        End If
        sprOrder.Col = 8: sprOrder.Text = adoSet.Fields("CollDate").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
    
Get_Patient_Data:
    spriOrder.Row = Row
    spriOrder.Col = 7
    
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INX_PATIENT0)  */"
    
    strSql = strSql & ""
    strSql = strSql & " SELECT a.*, b.Sname, b.Jumin1, b.Jumin2, b.Juso, b.Tel,"
    strSql = strSql & "        c.Postname1, c.Postname2, c.Postname3,"
    strSql = strSql & "        d.DeptnameK, e.Drname, f.WardCode"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Order   a,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PATIENT  b,"
    strSql = strSql & "        TW_MIS_PMPA.TWBas_Post     c,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT     d,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR   e,"
    strSql = strSql & "        TW_MIS_PMPA.TWBas_Room     f "
    strSql = strSql & " WHERE  a.Ptno = '" & spriOrder.Text & "'"
    strSql = strSql & " AND    a.Ptno = b.Ptno(+) "
    strSql = strSql & " AND    b.PostCode1 = c.PostCode1(+)"
    strSql = strSql & " AND    b.PostCode2 = c.PostCode2(+)"
    strSql = strSql & " AND    a.DeptCode  = d.DeptCode(+)"
    strSql = strSql & " AND    a.DrCode    = e.Drcode(+)"
    strSql = strSql & " AND    a.RoomCode  = f.RoomCode(+)"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    txtQryPtno.Text = adoSet.Fields("Ptno").Value & ""
    txtWard.Text = Trim(adoSet.Fields("WardCode").Value & "") & "/" & adoSet.Fields("RoomCode").Value & ""
    txtSname.Text = adoSet.Fields("Sname").Value & ""
    txtSexage.Text = adoSet.Fields("Sex").Value & "/" & adoSet.Fields("AgeYY").Value & ""
    txtJumin.Text = adoSet.Fields("JUmin1").Value & "-" & adoSet.Fields("Jumin2").Value & ""
    txtTel.Text = adoSet.Fields("Tel").Value & ""
    txtDept.Text = adoSet.Fields("DeptnameK").Value & ""
    txtDr.Text = adoSet.Fields("Drname").Value & ""
    txtAddr.Text = adoSet.Fields("PostName1").Value & " " & _
                   adoSet.Fields("PostName2").Value & " " & _
                   adoSet.Fields("PostName3").Value & " " & _
                   adoSet.Fields("Juso").Value & ""
    
    Call adoSetClose(adoSet)
    Return
    

End Sub

Private Sub sproOrder_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
    If Col = 1 Then
        sproOrder.Row = Row
        sproOrder.Col = 4: GLabelJeobsuDt = sproOrder.Text
        sproOrder.Col = 8: GLabelPtno = sproOrder.Text
        sproOrder.Col = 4: GLabelJDt = sproOrder.Text
        sproOrder.Col = 5: GLabelJT1 = sproOrder.Text
        sproOrder.Col = 6: GLabelJT2 = sproOrder.Text
        GLabelLoadCheck = "LOAD"
        frmBarCode.Show vbModal
    End If

End Sub

Private Sub sproOrder_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim sQryPtno        As String
    Dim sQryJeobsuDt    As String
    Dim sQryJeobsuT1    As String
    Dim sQryJeobsuT2    As String
    Dim sQryDeptC       As String
    
    Dim sGbio       As String
    Dim sVerify     As String
    Dim sFrDate     As String
    Dim sToDate     As String
        
    sFrDate = Format(dtFrDate.Value, "yyyy-MM-dd")
    sToDate = Format(dtToDate.Value, "yyyy-MM-dd")
        
        
        
    If Row = 0 Then Exit Sub
    If Row > sproOrder.DataRowCnt Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    sproOrder.Row = Row
    sproOrder.Col = 8: sQryPtno = sproOrder.Text
    sproOrder.Col = 2: sQryDeptC = sproOrder.Text
    sproOrder.Col = 4: sQryJeobsuDt = sproOrder.Text
    sproOrder.Col = 5: sQryJeobsuT1 = sproOrder.Text
    sproOrder.Col = 6: sQryJeobsuT2 = sproOrder.Text
        
    Call Spread_Set_Clear(sprOrder)
    GoSub Set_WhereData
    GoSub Get_General_Order
    GoSub Get_Patient_Data
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
    
Set_WhereData:
    If optOpd.Value = True Then
        sGbio = "O"
        sVerify = "Y"
    End If
    
    If optIpd.Value = True Then
        sGbio = "I"
        If optVer.Value = True Then sVerify = "Y"
        If optNoVer.Value = True Then sVerify = "12"
    End If
    
    Return
    
    
Get_General_Order:
    strSql = ""
    strSql = strSql & "  SELECT a.ItemCd Codeky, c.ItemNM ItemName, a.GeomchCd, d.Codenm Samplename, a.Orderno,"
    strSql = strSql & "         TO_CHAR(a.JeobsuDt,'yyyy-MM-dd') JeobsuDt,"
    strSql = strSql & "         TO_CHAR(b.GBDate, 'yyyy-MM-dd hh24:mi') GBDate, e.Name"
    strSql = strSql & "  FROM   TWEXAM_General_Sub a,        "
    strSql = strSql & "         TWEXAM_General     b,        "
    strSql = strSql & "         TWEXAM_itemML      c,        "
    strSql = strSql & "         TWEXAM_Sample      d,        "
    strSql = strSql & "         TW_MIS_PMPA.TWBas_Pass  e,   "
    strSql = strSql & "         TWEXAM_Order       f         "
    strSql = strSql & "  where  a.JeobsuDt  = TO_DATE('" & sQryJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & "  AND    a.Ptno      = '" & sQryPtno & "'"
    strSql = strSql & "  AND    b.JeobsuT1  = '" & sQryJeobsuT1 & "'"
    strSql = strSql & "  AND    b.JeobsuT2  = '" & sQryJeobsuT2 & "'"
    strSql = strSql & "  AND    a.Routincd  = a.itemCd "
    strSql = strSql & "  AND    a.JeobsuDt  = b.JeobsuDt(+) "
    strSql = strSql & "  AND    a.SLipno1   = b.SLipno1(+) "
    strSql = strSql & "  AND    a.SLipno2   = b.SLipno2(+) "
    strSql = strSql & "  AND    a.ItemCd    = c.Codeky"
    strSql = strSql & "  AND    a.JeobsuDt  = f.CollDate(+)"
    strSql = strSql & "  AND    a.Orderno   = f.Orderno(+)"
    strSql = strSql & "  AND    a.GeomchCd  = d.Code(+)"
    strSql = strSql & "  AND    b.GbCH      = 'Y' "
    strSql = strSql & "  AND    b.DeptCode  = '" & sQryDeptC & "'"
    strSql = strSql & "  AND    f.CollID    = e.Idnumber(+) "
    strSql = strSql & "  AND   ( e.ProgramID IS NULL Or e.ProgramID = ' ')"
    strSql = strSql & "  UNION ALL"
    strSql = strSql & "  SELECT distinct a.Routincd Codeky, c.Routinnm ItemName, a.GeomchCd, d.Codenm Samplename, a.Orderno,"
    strSql = strSql & "         TO_CHAR(a.JeobsuDt,'yyyy-MM-dd') JeobsuDt,"
    strSql = strSql & "         TO_CHAR(b.GBDate, 'yyyy-MM-dd hh24:mi') GBDate, e.Name"
    strSql = strSql & "  FROM   TWEXAM_General_Sub a,        "
    strSql = strSql & "         TWEXAM_General     b,        "
    strSql = strSql & "         TWEXAM_Routine     c,        "
    strSql = strSql & "         TWEXAM_Sample      d,        "
    strSql = strSql & "         TW_MIS_PMPA.TWBas_Pass e,    "
    strSql = strSql & "         TWEXAM_Order       f         "
    strSql = strSql & "  where  a.JeobsuDt  = TO_DATE('" & sQryJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & "  AND    a.Ptno      = '" & sQryPtno & "'"
    strSql = strSql & "  AND    b.JeobsuT1  = '" & sQryJeobsuT1 & "'"
    strSql = strSql & "  AND    b.JeobsuT2  = '" & sQryJeobsuT2 & "'"
    strSql = strSql & "  AND    a.Routincd != a.itemCd "
    strSql = strSql & "  AND    a.JeobsuDt  = b.JeobsuDt(+) "
    strSql = strSql & "  AND    a.SLipno1   = b.SLipno1(+) "
    strSql = strSql & "  AND    a.SLipno2   = b.SLipno2(+) "
    strSql = strSql & "  AND    a.RoutinCD  = c.RoutinCd"
    strSql = strSql & "  AND    a.GeomchCd  = d.Code(+)"
    strSql = strSql & "  AND    a.JeobsuDt  = f.CollDate(+)"
    strSql = strSql & "  AND    a.Orderno   = f.Orderno(+)"
    strSql = strSql & "  AND    b.GbCH      = 'Y' "
    strSql = strSql & "  AND    b.DeptCode  = '" & sQryDeptC & "'"
    strSql = strSql & "  AND    f.CollID    = e.Idnumber(+) "
    strSql = strSql & "  AND   ( e.ProgramID IS NULL Or e.ProgramID = ' ')"
    strSql = strSql & "  order by Codeky"
    
    If False = adoSetOpen(strSql, adoSet) Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Do Until adoSet.EOF
        sprOrder.Row = sprOrder.DataRowCnt + 1
        sprOrder.Col = 1: sprOrder.Text = Left(adoSet.Fields("Codeky").Value & "", 2)
        sprOrder.Col = 2: sprOrder.Text = adoSet.Fields("Codeky").Value & ""
        sprOrder.Col = 3: sprOrder.Text = adoSet.Fields("ItemName").Value & ""
        sprOrder.Col = 4: sprOrder.Text = adoSet.Fields("Samplename").Value & ""
        If sVerify = "Y" Then
            sprOrder.Col = 5: sprOrder.Text = adoSet.Fields("Name").Value & ""
            sprOrder.Col = 6: sprOrder.Text = adoSet.Fields("GbDate").Value & ""
            'sprOrder.Col = 6: sprOrder.Text = adoSet.Fields("GeomsaDt").Value & " " & _
            '                                  Format(adoSet.Fields("GeomsaT1").Value & "", "00") & ":" & _
            '                                  Format(adoSet.Fields("GeomsaT2").Value & "", "00")
            sprOrder.Col = 7: sprOrder.Text = adoSet.Fields("Orderno").Value & ""
        End If
        sprOrder.Col = 8: sprOrder.Text = adoSet.Fields("JeobsuDt").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return



Get_General_Order1:
    strSql = ""
    strSql = strSql & "   SELECT a.GeomChCd, c.Codenm SampleName, a.Orderno,"
    strSql = strSql & "          a.ItemCd,   b.ItemNM ItemName,"
    strSql = strSql & "          TO_CHAR(a.JeobsuDt,'YYYY-MM-DD') JeobsuDt,"
    strSql = strSql & "          a.JeobsuT1, a.JeobsuT2, a.Geomsaja, d.Name,"
    strSql = strSql & "          TO_CHAR(a.GeomsaDt,'YYYY-MM-DD') GeomsaDt,"
    strSql = strSql & "          GeomsaT1, GeomsaT2"
    strSql = strSql & "   FROM   TWEXAM_General  a,"
    strSql = strSql & "          TW_MIS_EXAM.TWEXAM_itemML  b,"
    strSql = strSql & "          TW_MIS_EXAM.TWEXAM_Sample  c,"
    strSql = strSql & "          TW_MIS_PMPA.TWBas_Pass     d "
    strSql = strSql & "   WHERE  a.Ptno     =  '" & sQryPtno & "'"
    strSql = strSql & "   AND    a.JeobsuDt = TO_DATE('" & sQryJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & "   AND    a.JeobsuT1 = '" & sQryJeobsuT1 & "'"
    strSql = strSql & "   AND    a.JeobsuT2 = '" & sQryJeobsuT2 & "'"
    strSql = strSql & "   AND    a.DeptCode = '" & sQryDeptC & "'"
    If sVerify = "Y" Then
        strSql = strSql & " AND  a.GbCH    = 'Y'"
    Else
        strSql = strSql & " AND  a.GbCh   IN ('1','2')"
    End If
    strSql = strSql & "   AND    a.ItemCd   = b.Codeky"
    strSql = strSql & "   AND    a.GeomChCd = c.Code(+)"
    strSql = strSql & "   AND    a.Geomsaja = d.Idnumber(+)"
    strSql = strSql & "   AND   ( d.ProgramID IS NULL Or d.ProgramID = ' ')"
    strSql = strSql & " Union ALL"
    strSql = strSql & "   SELECT DISTINCT a.GeomChCd,  c.Codenm SampleName, a.Orderno,"
    strSql = strSql & "                   a.ItemCd,    b.RoutinNM ItemName,"
    strSql = strSql & "          TO_CHAR(a.JeobsuDt,'YYYY-MM-DD') JeobsuDt,"
    strSql = strSql & "          a.JeobsuT1, a.JeobsuT2, a.Geomsaja, d.Name,"
    strSql = strSql & "          TO_CHAR(a.GeomsaDt,'YYYY-MM-DD') GeomsaDt,"
    strSql = strSql & "          GeomsaT1, GeomsaT2"
    strSql = strSql & "   FROM   TWEXAM_General a,"
    strSql = strSql & "          TW_MIS_EXAM.TWEXAM_Routine b,"
    strSql = strSql & "          TW_MIS_EXAM.TWEXAM_Sample  c,"
    strSql = strSql & "          TW_MIS_PMPA.TWBas_Pass     d "
    strSql = strSql & "   WHERE  a.Ptno     =  '" & sQryPtno & "'"
    strSql = strSql & "   AND    a.JeobsuDt = TO_DATE('" & sQryJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & "   AND    a.JeobsuT1 = '" & sQryJeobsuT1 & "'"
    strSql = strSql & "   AND    a.JeobsuT2 = '" & sQryJeobsuT2 & "'"
    strSql = strSql & "   AND    a.DeptCode = '" & sQryDeptC & "'"
    If sVerify = "Y" Then
        strSql = strSql & " AND  a.GbCH    = 'Y'"
    Else
        strSql = strSql & " AND  a.GbCh   IN ('1','2')"
    End If
    strSql = strSql & "   AND    a.ItemCd   = b.RoutinCd"
    strSql = strSql & "   AND    a.GeomChCd = c.Code(+)"
    strSql = strSql & "   AND    a.Geomsaja   = d.Idnumber(+)"
    strSql = strSql & "   AND   ( d.ProgramID IS NULL Or d.ProgramID = ' ')"
    strSql = strSql & "   ORDER  BY ItemCd"
    
    If False = adoSetOpen(strSql, adoSet) Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Do Until adoSet.EOF
        sprOrder.Row = sprOrder.DataRowCnt + 1
        sprOrder.Col = 1: sprOrder.Text = Left(adoSet.Fields("ItemCd").Value & "", 2)
        sprOrder.Col = 2: sprOrder.Text = adoSet.Fields("ItemCD").Value & ""
        sprOrder.Col = 3: sprOrder.Text = adoSet.Fields("ItemName").Value & ""
        sprOrder.Col = 4: sprOrder.Text = adoSet.Fields("Samplename").Value & ""
        If sVerify = "Y" Then
            sprOrder.Col = 5: sprOrder.Text = adoSet.Fields("Name").Value & ""
            sprOrder.Col = 6: sprOrder.Text = adoSet.Fields("GeomsaDt").Value & " " & _
                                              Format(adoSet.Fields("GeomsaT1").Value & "", "00") & ":" & _
                                              Format(adoSet.Fields("GeomsaT2").Value & "", "00")
            sprOrder.Col = 7: sprOrder.Text = adoSet.Fields("Orderno").Value & ""
        End If
        sprOrder.Col = 8: sprOrder.Text = adoSet.Fields("JeobsuDt").Value & ""
        adoSet.MoveNext
    Loop
    Call adoSetClose(adoSet)
    
    Return
    


Get_Patient_Data:
    sproOrder.Row = Row
    sproOrder.Col = 8
    
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INX_PATIENT0)  */"
    
    strSql = ""
    strSql = strSql & " SELECT a.*, b.Sname, b.Jumin1, b.Jumin2, b.Juso, b.Tel,"
    strSql = strSql & "        c.Postname1, c.Postname2, c.Postname3,"
    strSql = strSql & "        d.DeptnameK, e.Drname, f.WardCode"
    strSql = strSql & " FROM   TW_MIS_EXAM.TWEXAM_Order   a,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PATIENT  b,"
    strSql = strSql & "        TW_MIS_PMPA.TWBas_Post     c,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT     d,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR   e,"
    strSql = strSql & "        TW_MIS_PMPA.TWBas_Room     f "
    strSql = strSql & " WHERE  a.Ptno      = '" & sproOrder.Text & "'"
    strSql = strSql & " AND    a.COLLDate  = TO_DATE('" & sQryJeobsuDt & "','YYYY-MM-DD')"
    strSql = strSql & " AND    a.COLLHH  = '" & sQryJeobsuT1 & "'"
    strSql = strSql & " AND    a.COLLMM  = '" & sQryJeobsuT2 & "'"
    strSql = strSql & " AND    a.GbCH      = 'Y'"
    strSql = strSql & " AND    a.Ptno      = b.Ptno(+) "
    strSql = strSql & " AND    b.PostCode1 = c.PostCode1(+)"
    strSql = strSql & " AND    b.PostCode2 = c.PostCode2(+)"
    strSql = strSql & " AND    a.DeptCode  = d.DeptCode(+)"
    strSql = strSql & " AND    a.DrCode    = e.Drcode(+)"
    strSql = strSql & " AND    a.RoomCode  = f.RoomCode(+)"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    txtQryPtno.Text = adoSet.Fields("Ptno").Value & ""
    txtWard.Text = ""
    txtSname.Text = adoSet.Fields("Sname").Value & ""
    txtSexage.Text = adoSet.Fields("Sex").Value & "/" & adoSet.Fields("AgeYY").Value & ""
    txtJumin.Text = adoSet.Fields("JUmin1").Value & "-" & adoSet.Fields("Jumin2").Value & ""
    txtTel.Text = adoSet.Fields("Tel").Value & ""
    txtDept.Text = adoSet.Fields("DeptnameK").Value & ""
    txtDr.Text = adoSet.Fields("Drname").Value & ""
    txtAddr.Text = adoSet.Fields("PostName1").Value & " " & _
                   adoSet.Fields("PostName2").Value & " " & _
                   adoSet.Fields("PostName3").Value & " " & _
                   adoSet.Fields("Juso").Value & ""
    
    Call adoSetClose(adoSet)
    Return
    
End Sub


Private Sub sprOrder_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    
    frmQryRet.Show vbModal
    
    
End Sub

Private Sub txtQryPtno_KeyPress(KeyAscii As Integer)
    
    
    
    If KeyAscii = 13 Then
        If Trim(txtQryPtno.Text) = "" Then Exit Sub
        If Len(Trim(txtQryPtno.Text)) = 12 Then
            txtQryPtno.Text = convLabnoToPtno(txtQryPtno.Text)
        End If
        If Trim(txtQryPtno.Text) = "" Then Exit Sub
        
        sFlagPtnoBoxEnter = "13"
        txtQryPtno.Text = Format(txtQryPtno.Text, "00000000")
        Call Spread_Set_Clear(sprOrder)
        GoSub Right_TextBox_Clear
        GoSub Patient_Check_Sub
        
        If optOpd.Value = True Then
            Call cmdoOrder_Click
            If sproOrder.DataRowCnt > 0 Then
                Call sproOrder_DblClick(2, 1)
            End If
        End If
        If optIpd.Value = True Then
            Call cmdiOrder_Click
            If spriOrder.DataRowCnt > 0 Then
                Call spriOrder_DblClick(2, 1)
            End If
            
        End If
        sFlagPtnoBoxEnter = ""
    End If
    
    Exit Sub

'/-----------------------------------------------------------
Patient_Check_Sub:
    'strSql = ""
    'strSql = strSql & " SELECT /*+ INDEX (TW_MIS_PMPA.TWBAS_PATIENT INX_PATIENT0)  */"
    
    strSql = ""
    strSql = strSql & " SELECT a.*, b.Jumin1, b.Jumin2, b.Juso, b.Tel,"
    strSql = strSql & "        c.Postname1, c.Postname2, c.Postname3,"
    strSql = strSql & "        d.DeptnameK, e.Drname, f.WardCode"
    strSql = strSql & " FROM   TWEXAM_IDNOMST a,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_PATIENT  b,"
    strSql = strSql & "        TW_MIS_PMPA.TWBas_Post     c,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DEPT     d,"
    strSql = strSql & "        TW_MIS_PMPA.TWBAS_DOCTOR   e,"
    strSql = strSql & "        TW_MIS_PMPA.TWBas_Room     f "
    strSql = strSql & " WHERE  a.Ptno = '" & txtQryPtno.Text & "'"
    strSql = strSql & " AND    a.Ptno = b.Ptno(+) "
    strSql = strSql & " AND    b.PostCode1 = c.PostCode1(+)"
    strSql = strSql & " AND    b.PostCode2 = c.PostCode2(+)"
    strSql = strSql & " AND    a.DeptCode  = d.DeptCode(+)"
    strSql = strSql & " AND    a.DrCode    = e.Drcode(+)"
    strSql = strSql & " AND    a.RoomCode  = f.RoomCode(+)"
    
    If False = adoSetOpen(strSql, adoSet) Then Return
    
    txtWard.Text = Trim(adoSet.Fields("WardCode").Value & "") & "/" & adoSet.Fields("RoomCode").Value & ""
    txtSname.Text = adoSet.Fields("Sname").Value & ""
    txtSexage.Text = adoSet.Fields("Sex").Value & "/" & adoSet.Fields("AgeYY").Value & ""
    txtJumin.Text = adoSet.Fields("JUmin1").Value & "-" & adoSet.Fields("Jumin2").Value & ""
    txtTel.Text = adoSet.Fields("Tel").Value & ""
    txtDept.Text = adoSet.Fields("DeptnameK").Value & ""
    txtDr.Text = adoSet.Fields("Drname").Value & ""
    txtAddr.Text = adoSet.Fields("PostName1").Value & " " & _
                   adoSet.Fields("PostName2").Value & " " & _
                   adoSet.Fields("PostName3").Value & " " & _
                   adoSet.Fields("Juso").Value & ""
    
    Call adoSetClose(adoSet)
    
    Return
    
    
Right_TextBox_Clear:
    
    txtWard.Text = ""
    txtSname.Text = ""
    txtSexage.Text = ""
    txtJumin.Text = ""
    txtTel.Text = ""
    txtDept.Text = ""
    txtDr.Text = ""
    txtAddr.Text = ""
    Return
    
End Sub
