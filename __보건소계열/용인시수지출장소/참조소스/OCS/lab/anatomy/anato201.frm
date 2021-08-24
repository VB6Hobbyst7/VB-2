VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Anato_Jeobsu_View 
   Caption         =   "접수환자명단"
   ClientHeight    =   7245
   ClientLeft      =   1650
   ClientTop       =   1905
   ClientWidth     =   8760
   Icon            =   "ANATO201.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7245
   ScaleWidth      =   8760
   Begin VB.Frame frmComplete 
      Caption         =   "결과완료자 조회"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   804
      Left            =   2688
      TabIndex        =   4
      Top             =   2232
      Visible         =   0   'False
      Width           =   4452
      Begin VB.TextBox txtClass 
         BackColor       =   &H00DCFAFA&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   192
         MaxLength       =   2
         TabIndex        =   0
         Top             =   270
         Width           =   600
      End
      Begin VB.TextBox txtDateYY 
         BackColor       =   &H00DCFAFA&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   912
         MaxLength       =   4
         TabIndex        =   1
         Top             =   270
         Width           =   600
      End
      Begin VB.TextBox txtSeqnum 
         BackColor       =   &H00DCFAFA&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1632
         MaxLength       =   5
         TabIndex        =   2
         Top             =   270
         Width           =   1400
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   480
         Left            =   3288
         TabIndex        =   3
         Top             =   204
         Width           =   996
         _Version        =   65536
         _ExtentX        =   1757
         _ExtentY        =   847
         _StockProps     =   78
         Caption         =   "조 회"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         Font3D          =   1
         RoundedCorners  =   0   'False
         AutoSize        =   1
      End
   End
   Begin Threed.SSCommand cmdResult 
      Height          =   756
      Left            =   7296
      TabIndex        =   17
      Top             =   4872
      Width           =   1284
      _Version        =   65536
      _ExtentX        =   2265
      _ExtentY        =   1333
      _StockProps     =   78
      Caption         =   "결과조회"
      Picture         =   "ANATO201.frx":030A
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   756
      Left            =   7296
      TabIndex        =   15
      Top             =   4020
      Width           =   1284
      _Version        =   65536
      _ExtentX        =   2265
      _ExtentY        =   1333
      _StockProps     =   78
      Caption         =   "종    료"
      Picture         =   "ANATO201.frx":0624
   End
   Begin Threed.SSCommand cmdSelect 
      Height          =   756
      Left            =   7296
      TabIndex        =   14
      Top             =   3180
      Width           =   1284
      _Version        =   65536
      _ExtentX        =   2265
      _ExtentY        =   1333
      _StockProps     =   78
      Caption         =   "조    회"
      Picture         =   "ANATO201.frx":093E
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   2010
      Left            =   3945
      TabIndex        =   9
      Top             =   120
      Width           =   2700
      _Version        =   65536
      _ExtentX        =   4762
      _ExtentY        =   3545
      _StockProps     =   14
      Caption         =   "입력선택"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSOption optAdditional 
         Height          =   255
         Left            =   195
         TabIndex        =   22
         Top             =   1650
         Width           =   2385
         _Version        =   65536
         _ExtentX        =   4207
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Additional"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optPreliminary 
         Height          =   255
         Left            =   195
         TabIndex        =   21
         Top             =   975
         Width           =   2385
         _Version        =   65536
         _ExtentX        =   4207
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "판독입력및 판독수정"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optDiag 
         Height          =   255
         Left            =   195
         TabIndex        =   20
         Top             =   1320
         Width           =   2385
         _Version        =   65536
         _ExtentX        =   4207
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "결과완료입력"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optEyeCheck 
         Height          =   255
         Left            =   195
         TabIndex        =   19
         Top             =   645
         Width           =   2385
         _Version        =   65536
         _ExtentX        =   4207
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Preliminary입력"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optJeobsu 
         Height          =   255
         Left            =   195
         TabIndex        =   18
         Top             =   315
         Width           =   2385
         _Version        =   65536
         _ExtentX        =   4207
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "육안결과입력및 수정"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   2010
      Left            =   2070
      TabIndex        =   8
      Top             =   120
      Width           =   1650
      _Version        =   65536
      _ExtentX        =   2900
      _ExtentY        =   3535
      _StockProps     =   14
      Caption         =   "조회방법"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   8.98
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSOption optRefferal 
         Height          =   255
         Left            =   210
         TabIndex        =   23
         Top             =   1110
         Width           =   1260
         _Version        =   65536
         _ExtentX        =   2222
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Refferal"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optAll 
         Height          =   255
         Left            =   210
         TabIndex        =   16
         Top             =   1485
         Width           =   1260
         _Version        =   65536
         _ExtentX        =   2222
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "ALL"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin Threed.SSOption optCytology 
         Height          =   255
         Left            =   210
         TabIndex        =   13
         Top             =   735
         Width           =   1260
         _Version        =   65536
         _ExtentX        =   2222
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Cytology"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optHistology 
         Height          =   255
         Left            =   210
         TabIndex        =   12
         Top             =   360
         Width           =   1260
         _Version        =   65536
         _ExtentX        =   2222
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Histology"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.01
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2010
      Left            =   195
      TabIndex        =   7
      Top             =   120
      Width           =   1590
      _Version        =   65536
      _ExtentX        =   2805
      _ExtentY        =   3545
      _StockProps     =   14
      Caption         =   "정렬방법"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSOption optSname 
         Height          =   360
         Left            =   195
         TabIndex        =   11
         Top             =   1425
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   635
         _StockProps     =   78
         Caption         =   "환자명"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optSeqnum 
         Height          =   360
         Left            =   195
         TabIndex        =   10
         Top             =   405
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   635
         _StockProps     =   78
         Caption         =   "병리번호"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSCommand cmdOk 
      Height          =   756
      Left            =   7296
      TabIndex        =   6
      Top             =   2328
      Width           =   1284
      _Version        =   65536
      _ExtentX        =   2265
      _ExtentY        =   1333
      _StockProps     =   78
      Caption         =   "선택완료"
      Picture         =   "ANATO201.frx":0D90
   End
   Begin FPSpread.vaSpread ssRecept 
      Height          =   4710
      Left            =   195
      TabIndex        =   5
      Top             =   2325
      Width           =   6945
      _Version        =   196608
      _ExtentX        =   12250
      _ExtentY        =   8308
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   0
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
      ScrollBars      =   2
      ShadowColor     =   12632256
      ShadowDark      =   8421504
      ShadowText      =   0
      SpreadDesigner  =   "ANATO201.frx":10AA
      UserResize      =   0
      VisibleCols     =   500
      VisibleRows     =   500
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1932
      Left            =   6888
      TabIndex        =   24
      Top             =   192
      Width           =   1692
      _Version        =   65536
      _ExtentX        =   2984
      _ExtentY        =   3408
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   8.99
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelInner      =   1
      Begin MSComCtl2.DTPicker dtToJeobsu 
         Height          =   315
         Left            =   240
         TabIndex        =   25
         Top             =   1380
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24444931
         CurrentDate     =   36312
      End
      Begin MSComCtl2.DTPicker dtFromJeobsu 
         Height          =   315
         Left            =   240
         TabIndex        =   26
         Top             =   780
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24444931
         CurrentDate     =   36312
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00808000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "검사완료일"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   90
         TabIndex        =   29
         Top             =   120
         Width           =   1485
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   28
         Top             =   510
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   27
         Top             =   1140
         Width           =   210
      End
   End
End
Attribute VB_Name = "Anato_Jeobsu_View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TabCheck                As Integer


Private Sub Form_Load()
    
    dtFromJeobsu.Value = Dual_Date_Get("yyyy-MM-dd")
    dtToJeobsu.Value = Dual_Date_Get("yyyy-MM-dd")
    
    SSPanel1.Visible = False
    
    frmComplete.Enabled = True
    
'    Anato_Jeobsu_View.Width = 6120      '6165
'    Anato_Jeobsu_View.Height = 7860     ' 8040
    
    optSeqnum.Value = True
    optHistology.Value = True
    optJeobsu.Value = True
    
    Call cmdSelect_Click

End Sub


Private Sub cmdExit_Click()
    GAnato_Jeobsu_View = False
    Unload Me

End Sub


Private Sub cmdResult_Click()
    optAdditional.Value = True
    frmComplete.Visible = True
    txtClass.SetFocus
    
End Sub


Private Sub cmdSelect_Click()
    Dim rs                  As ADODB.Recordset
    Dim i                   As Integer
    
    gSFrDate = Format(dtFromJeobsu.Value, "yyyy-MM-dd")
    gSToDate = Format(dtToJeobsu.Value, "yyyy-MM-dd")
    
    Call SSInitialize(ssRecept)
    
    strSQL = ""
    strSQL = strSQL & " SELECT a.*, a.RowID,"
    strSQL = strSQL & "        TO_CHAR(a.Jdate,   'YYYY-MM-DD') jdate,"
    strSQL = strSQL & "        TO_CHAR(a.Orderdt, 'YYYY-MM-DD') Orderdt"
    strSQL = strSQL & " FROM   TWANAT_DIAG a"
    strSQL = strSQL & " WHERE  GBRESULT <> 'X' "
    
    If optHistology.Value = True Then
        strSQL = strSQL & "   AND  CLASS  = 'P' "      'Histology
    ElseIf optCytology.Value = True Then
        strSQL = strSQL & "   AND  CLASS  = 'C' "      'cytology
    ElseIf optRefferal.Value = True Then
        strSQL = strSQL & "   AND  CLASS  = 'R' "      'REFFERAL
    End If
    
    If optJeobsu.Value = True Then                      '접수완료
        strSQL = strSQL & " AND  GBRESULT = '0' "
        strSQL = strSQL & " AND  GBGROSS  <= '1' "
    ElseIf optEyeCheck.Value = True Then                '육안결과 Gross
'        strSQL = strSQL & " AND  GBGROSS  = '1' "
        strSQL = strSQL & " AND  GBGROSS  <= '1' "
        strSQL = strSQL & " AND  GBRESULT = '0' "
    ElseIf optPreliminary.Value = True Then             'Preliminary
        strSQL = strSQL & " AND  ( GBRESULT = '0' OR GBRESULT = '2' OR  GBRESULT = '3' ) "
    ElseIf optDiag.Value = True Then                    '판독
        strSQL = strSQL & " AND  GBRESULT = '3' "
    ElseIf optAdditional.Value = True Then              'Additional
        strSQL = strSQL & " AND  GBRESULT >=  '4' "
        strSQL = strSQL & " AND  a.DiagDate   BETWEEN TO_DATE('" & gSFrDate & "','YYYY-MM-DD')"
        strSQL = strSQL & "                       AND TO_DATE('" & gSToDate & "','YYYY-MM-DD')"
    End If
    
    If optSeqnum.Value = True Then
        strSQL = strSQL & " ORDER  BY CLASS, DATEYY, SEQNUM ASC "
    Else
        strSQL = strSQL & " ORDER  BY SNAME ASC, CLASS, DATEYY, SEQNUM ASC "
    End If
    
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
'    ssRecept.MaxRows = Rowindicator + 1
    
    Do Until rs.EOF
        ssRecept.Row = ssRecept.DataRowCnt + 1
        
        ssRecept.Row = i + 1
        ssRecept.Col = 2:  ssRecept.Text = rs.Fields("Class").Value & "-" & _
                                           rs.Fields("Dateyy").Value & "-" & _
                                           rs.Fields("Seqnum").Value & ""
        ssRecept.Col = 3:  ssRecept.Text = rs.Fields("Ptno").Value & ""
        ssRecept.Col = 4:  ssRecept.Text = rs.Fields("Sname").Value & ""
        
        If rs.Fields("GbGross").Value & "" = "1" Then
            ssRecept.Col = 5:    ssRecept.Text = "육안결과"
        End If
        
        Select Case rs.Fields("GbResult").Value & ""
            Case "0"
                    If rs.Fields("GbGross").Value & "" = "1" Then
                        ssRecept.Col = 5: ssRecept.Text = "육안결과"
                    Else
                        ssRecept.Col = 5: ssRecept.Text = "접수중"
                    End If
            Case "2"
                    ssRecept.Col = 5:    ssRecept.Text = "Preliminary"
            Case "3"
                    ssRecept.Col = 5:    ssRecept.Text = "판독"
            Case "4"
                    ssRecept.Col = 5:    ssRecept.Text = "결과완료"
            Case "9"
                    ssRecept.Col = 5:    ssRecept.Text = "Additional"
            Case Else
        End Select
        
        ssRecept.Col = 6:   ssRecept.Text = rs.Fields("sex").Value & ""
        ssRecept.Col = 7:   ssRecept.Text = rs.Fields("ageyy").Value & ""
        
        ssRecept.Col = 8:   ssRecept.Text = rs.Fields("OrderDt").Value & ""
        ssRecept.Col = 9:   ssRecept.Text = rs.Fields("GbGross").Value & ""
        ssRecept.Col = 10:  ssRecept.Text = rs.Fields("GbResult").Value & ""
        ssRecept.Col = 11:  ssRecept.Text = rs.Fields("jdate").Value & ""
        
        ssRecept.Col = 12:  ssRecept.Text = rs.Fields("slid").Value & ""
        ssRecept.Col = 13:  ssRecept.Text = rs.Fields("PHOTO").Value & ""
        
        
        ssRecept.Col = 18:  ssRecept.Text = rs.Fields("ElectroScope").Value & ""
        ssRecept.Col = 19:  ssRecept.Text = rs.Fields("Flow").Value & ""
        
        ssRecept.Col = 20:  ssRecept.Text = rs.Fields("RowID").Value & ""
        
        ssRecept.Col = 21:  ssRecept.Text = rs.Fields("Special01").Value & ""
        ssRecept.Col = 22:  ssRecept.Text = rs.Fields("Special02").Value & ""
        ssRecept.Col = 23:  ssRecept.Text = rs.Fields("Special03").Value & ""
        ssRecept.Col = 24:  ssRecept.Text = rs.Fields("Special04").Value & ""
        ssRecept.Col = 25:  ssRecept.Text = rs.Fields("Special05").Value & ""
        ssRecept.Col = 26:  ssRecept.Text = rs.Fields("Special06").Value & ""
        ssRecept.Col = 27:  ssRecept.Text = rs.Fields("Special07").Value & ""
        ssRecept.Col = 28:  ssRecept.Text = rs.Fields("Special08").Value & ""
        ssRecept.Col = 29:  ssRecept.Text = rs.Fields("Special09").Value & ""
        ssRecept.Col = 30:  ssRecept.Text = rs.Fields("Special10").Value & ""
        
        ssRecept.Col = 31:  ssRecept.Text = rs.Fields("Special11").Value & ""
        ssRecept.Col = 32:  ssRecept.Text = rs.Fields("Special12").Value & ""
        ssRecept.Col = 33:  ssRecept.Text = rs.Fields("Special13").Value & ""
        ssRecept.Col = 34:  ssRecept.Text = rs.Fields("Special14").Value & ""
        ssRecept.Col = 35:  ssRecept.Text = rs.Fields("Special15").Value & ""
        ssRecept.Col = 36:  ssRecept.Text = rs.Fields("Special16").Value & ""
        ssRecept.Col = 37:  ssRecept.Text = rs.Fields("Special17").Value & ""
        ssRecept.Col = 38:  ssRecept.Text = rs.Fields("Special18").Value & ""
        ssRecept.Col = 39:  ssRecept.Text = rs.Fields("Special19").Value & ""
        ssRecept.Col = 40:  ssRecept.Text = rs.Fields("Special20").Value & ""
        
        ssRecept.Col = 41:  ssRecept.Text = rs.Fields("Special21").Value & ""
        ssRecept.Col = 42:  ssRecept.Text = rs.Fields("Special22").Value & ""
        ssRecept.Col = 43:  ssRecept.Text = rs.Fields("Special23").Value & ""
        ssRecept.Col = 44:  ssRecept.Text = rs.Fields("Special24").Value & ""
        ssRecept.Col = 45:  ssRecept.Text = rs.Fields("Special25").Value & ""
        ssRecept.Col = 46:  ssRecept.Text = rs.Fields("Special26").Value & ""
        ssRecept.Col = 47:  ssRecept.Text = rs.Fields("Special27").Value & ""
        ssRecept.Col = 48:  ssRecept.Text = rs.Fields("Special28").Value & ""
        ssRecept.Col = 49:  ssRecept.Text = rs.Fields("Special29").Value & ""
        ssRecept.Col = 50:  ssRecept.Text = rs.Fields("Special30").Value & ""
        
        rs.MoveNext
    
        i = i + 1
    Loop
    AdoCloseSet rs

End Sub


Private Sub cmdSearch_Click()
    '조회(결과완료)
    Dim rs                  As ADODB.Recordset
    
    frmComplete.Enabled = True
    Call SSInitialize(ssRecept)
    
    strSQL = ""
    strSQL = strSQL & " SELECT a.*, a.RowID,"
    strSQL = strSQL & "        TO_CHAR(a.Jdate,   'YYYY-MM-DD') jdate,"
    strSQL = strSQL & "        TO_CHAR(a.Orderdt, 'YYYY-MM-DD') Orderdt"
    strSQL = strSQL & " FROM   TWANAT_DIAG a"
    strSQL = strSQL & " WHERE  a.GbResult >=  '4'"
    strSQL = strSQL & " AND    a.Class    = '" & txtClass & "'  "
    strSQL = strSQL & " AND    a.DateYY   = '" & txtDateYY & "' "
    strSQL = strSQL & " AND    a.Seqnum   = '" & txtSeqnum & "' "
    strSQL = strSQL & " ORDER  BY Class, Dateyy, Seqnum "
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then frmComplete.Visible = False: Exit Sub
    
    Do Until rs.EOF
        ssRecept.Row = ssRecept.DataRowCnt + 1
        ssRecept.Col = 2:  ssRecept.Text = rs.Fields("Class").Value & "-" & _
                                           rs.Fields("Dateyy").Value & "-" & _
                                           rs.Fields("Seqnum").Value & ""
        ssRecept.Col = 3:  ssRecept.Text = rs.Fields("Ptno").Value & ""
        ssRecept.Col = 4:  ssRecept.Text = rs.Fields("Sname").Value & ""
        If rs.Fields("GbResult").Value & "" = "4" Then
            ssRecept.Col = 5: ssRecept.Text = "결과완료"
        ElseIf rs.Fields("GbResult").Value & "" = "9" Then
            ssRecept.Col = 5: ssRecept.Text = "Additional"
        End If
        ssRecept.Col = 6:   ssRecept.Text = rs.Fields("sex").Value & ""
        ssRecept.Col = 7:   ssRecept.Text = rs.Fields("ageyy").Value & ""
        
        ssRecept.Col = 8:   ssRecept.Text = rs.Fields("OrderDt").Value & ""
        ssRecept.Col = 9:   ssRecept.Text = rs.Fields("GbGross").Value & ""
        ssRecept.Col = 10:  ssRecept.Text = rs.Fields("GbResult").Value & ""
        ssRecept.Col = 11:  ssRecept.Text = rs.Fields("jdate").Value & ""
        
        ssRecept.Col = 12:  ssRecept.Text = rs.Fields("slid").Value & ""
        ssRecept.Col = 13:  ssRecept.Text = rs.Fields("PHOTO").Value & ""
        
        
        ssRecept.Col = 18:  ssRecept.Text = rs.Fields("ElectroScope").Value & ""
        ssRecept.Col = 19:  ssRecept.Text = rs.Fields("Flow").Value & ""
        
        ssRecept.Col = 20:  ssRecept.Text = rs.Fields("RowID").Value & ""
        
        ssRecept.Col = 21:  ssRecept.Text = rs.Fields("Special01").Value & ""
        ssRecept.Col = 22:  ssRecept.Text = rs.Fields("Special02").Value & ""
        ssRecept.Col = 23:  ssRecept.Text = rs.Fields("Special03").Value & ""
        ssRecept.Col = 24:  ssRecept.Text = rs.Fields("Special04").Value & ""
        ssRecept.Col = 25:  ssRecept.Text = rs.Fields("Special05").Value & ""
        ssRecept.Col = 26:  ssRecept.Text = rs.Fields("Special06").Value & ""
        ssRecept.Col = 27:  ssRecept.Text = rs.Fields("Special07").Value & ""
        ssRecept.Col = 28:  ssRecept.Text = rs.Fields("Special08").Value & ""
        ssRecept.Col = 29:  ssRecept.Text = rs.Fields("Special09").Value & ""
        ssRecept.Col = 30:  ssRecept.Text = rs.Fields("Special10").Value & ""
        
        ssRecept.Col = 31:  ssRecept.Text = rs.Fields("Special11").Value & ""
        ssRecept.Col = 32:  ssRecept.Text = rs.Fields("Special12").Value & ""
        ssRecept.Col = 33:  ssRecept.Text = rs.Fields("Special13").Value & ""
        ssRecept.Col = 34:  ssRecept.Text = rs.Fields("Special14").Value & ""
        ssRecept.Col = 35:  ssRecept.Text = rs.Fields("Special15").Value & ""
        ssRecept.Col = 36:  ssRecept.Text = rs.Fields("Special16").Value & ""
        ssRecept.Col = 37:  ssRecept.Text = rs.Fields("Special17").Value & ""
        ssRecept.Col = 38:  ssRecept.Text = rs.Fields("Special18").Value & ""
        ssRecept.Col = 39:  ssRecept.Text = rs.Fields("Special19").Value & ""
        ssRecept.Col = 40:  ssRecept.Text = rs.Fields("Special20").Value & ""
        
        ssRecept.Col = 41:  ssRecept.Text = rs.Fields("Special21").Value & ""
        ssRecept.Col = 42:  ssRecept.Text = rs.Fields("Special22").Value & ""
        ssRecept.Col = 43:  ssRecept.Text = rs.Fields("Special23").Value & ""
        ssRecept.Col = 44:  ssRecept.Text = rs.Fields("Special24").Value & ""
        ssRecept.Col = 45:  ssRecept.Text = rs.Fields("Special25").Value & ""
        ssRecept.Col = 46:  ssRecept.Text = rs.Fields("Special26").Value & ""
        ssRecept.Col = 47:  ssRecept.Text = rs.Fields("Special27").Value & ""
        ssRecept.Col = 48:  ssRecept.Text = rs.Fields("Special28").Value & ""
        ssRecept.Col = 49:  ssRecept.Text = rs.Fields("Special29").Value & ""
        ssRecept.Col = 50:  ssRecept.Text = rs.Fields("Special30").Value & ""
        
        
        rs.MoveNext
    Loop
    AdoCloseSet rs
    
    frmComplete.Visible = False
    
    Anato_Result.txtDiag.Locked = True
    Anato_DiagName_Input.txtSummary.Locked = False

End Sub


Private Sub CmdOK_Click()
    '선택완료
    Dim i                   As Integer
    
    Dim FCheck              As Boolean
    
    For i = 1 To ssRecept.DataRowCnt
        ssRecept.Row = i
        ssRecept.Col = 1
        If ssRecept.Text = "1" Then
            FCheck = True
        End If
    Next i
    If FCheck = False Then Exit Sub
    
    Select Case optCytology.Value
            Case True
                    If optJeobsu.Value = True Then
                        Anato_Result.cmdEyeCheck.Enabled = False
                '        Anato_Result.cmdGross.Enabled = False
                        Anato_Result.cmdJCode.Enabled = True
                        Anato_Result.cmdPreliminary.Enabled = True
                        Anato_Result.cmdFirstDiag.Enabled = True
                        Anato_Result.cmdSignOut.Enabled = False 'True
                        Anato_Result.cmdAdditional.Enabled = False
                        Anato_Result.cmdJCode.Enabled = False 'True
                        Anato_Result.txtSlid.Enabled = False
                        Anato_Result.SpinButton1.Enabled = False
                        Anato_Result.frmPhoto.Enabled = False
                
'                    ElseIf optEyeCheck.Value = True Then
'                        Anato_Result.cmdEyeCheck.Enabled = False
'                '        Anato_Result.cmdGross.Enabled = False
'                        Anato_Result.cmdJCode.Enabled = True
'                        Anato_Result.cmdPreliminary.Enabled = True
'                        Anato_Result.cmdFirstDiag.Enabled = True
'                        Anato_Result.cmdSignOut.Enabled = False 'True
'                        Anato_Result.cmdAdditional.Enabled = False
'                        Anato_Result.cmdJCode.Enabled = False 'True
'                        Anato_Result.txtSlid.Enabled = False
'                        Anato_Result.SpinButton1.Enabled = False
'                        Anato_Result.frmPhoto.Enabled = False
                    
                    ElseIf optPreliminary.Value = True Then
                        Anato_Result.cmdEyeCheck.Enabled = False
                '        Anato_Result.cmdGross.Enabled = False
                        Anato_Result.cmdJCode.Enabled = True
                        Anato_Result.cmdPreliminary.Enabled = False 'True
                        Anato_Result.cmdFirstDiag.Enabled = True
                        Anato_Result.cmdSignOut.Enabled = False
                        Anato_Result.cmdAdditional.Enabled = False
                        Anato_Result.cmdJCode.Enabled = True
                        Anato_Result.txtSlid.Enabled = True
                        Anato_Result.SpinButton1.Enabled = True
                        Anato_Result.frmPhoto.Enabled = True
                    
                    ElseIf optDiag.Value = True Then
                        Anato_Result.cmdEyeCheck.Enabled = False
                '        Anato_Result.cmdGross.Enabled = False
                        Anato_Result.cmdJCode.Enabled = True
                        Anato_Result.cmdPreliminary.Enabled = False
                        Anato_Result.cmdFirstDiag.Enabled = False
                        Anato_Result.cmdSignOut.Enabled = True
                        Anato_Result.cmdAdditional.Enabled = False
                        Anato_Result.cmdJCode.Enabled = True
                        Anato_Result.txtSlid.Enabled = True
                        Anato_Result.SpinButton1.Enabled = True
                        Anato_Result.frmPhoto.Enabled = True
                    
                    ElseIf optAdditional.Value = True Then
                        Anato_Result.cmdEyeCheck.Enabled = False
                '        Anato_Result.cmdGross.Enabled = False
                        Anato_Result.cmdJCode.Enabled = True
                        Anato_Result.cmdPreliminary.Enabled = False
                        Anato_Result.cmdFirstDiag.Enabled = False
                        Anato_Result.cmdSignOut.Enabled = False
                        Anato_Result.cmdAdditional.Enabled = True
                        Anato_Result.cmdJCode.Enabled = False
                        Anato_Result.txtSlid.Enabled = False
                        Anato_Result.SpinButton1.Enabled = False
                        Anato_Result.frmPhoto.Enabled = False
                    End If
            Case False
                    If optJeobsu.Value = True Then
                        Anato_Result.cmdEyeCheck.Enabled = True         '육안검사
                '        Anato_Result.cmdGross.Enabled = True            'Gross
                        Anato_Result.cmdJCode.Enabled = False           '진단코드
                        Anato_Result.cmdPreliminary.Enabled = False     'Preliminary
                        Anato_Result.cmdFirstDiag.Enabled = False       '판독
                        Anato_Result.cmdSignOut.Enabled = False         '결과완료
                        Anato_Result.cmdAdditional.Enabled = False      '추가진단
                        Anato_Result.cmdJCode.Enabled = False
                        Anato_Result.txtSlid.Enabled = False
                        Anato_Result.SpinButton1.Enabled = False
                        Anato_Result.frmPhoto.Enabled = False
                
                    ElseIf optEyeCheck.Value = True Then
                        Anato_Result.cmdEyeCheck.Enabled = False
                '        Anato_Result.cmdGross.Enabled = False
                        Anato_Result.cmdJCode.Enabled = True
                        Anato_Result.cmdPreliminary.Enabled = True
                        Anato_Result.cmdFirstDiag.Enabled = True
                        Anato_Result.cmdSignOut.Enabled = False 'True
                        Anato_Result.cmdAdditional.Enabled = False
                        Anato_Result.cmdJCode.Enabled = False 'True
                        Anato_Result.txtSlid.Enabled = False
                        Anato_Result.SpinButton1.Enabled = False
                        Anato_Result.frmPhoto.Enabled = False
                    
                    ElseIf optPreliminary.Value = True Then
                        Anato_Result.cmdEyeCheck.Enabled = False
                '        Anato_Result.cmdGross.Enabled = False
                        Anato_Result.cmdJCode.Enabled = True
                        Anato_Result.cmdPreliminary.Enabled = False 'True
                        Anato_Result.cmdFirstDiag.Enabled = True
                        Anato_Result.cmdSignOut.Enabled = False
                        Anato_Result.cmdAdditional.Enabled = False
                        Anato_Result.cmdJCode.Enabled = True
                        Anato_Result.txtSlid.Enabled = True
                        Anato_Result.SpinButton1.Enabled = True
                        Anato_Result.frmPhoto.Enabled = True
                    
                    ElseIf optDiag.Value = True Then
                        Anato_Result.cmdEyeCheck.Enabled = False
                '        Anato_Result.cmdGross.Enabled = False
                        Anato_Result.cmdJCode.Enabled = True
                        Anato_Result.cmdPreliminary.Enabled = False
                        Anato_Result.cmdFirstDiag.Enabled = False
                        Anato_Result.cmdSignOut.Enabled = True
                        Anato_Result.cmdAdditional.Enabled = False
                        Anato_Result.cmdJCode.Enabled = True
                        Anato_Result.txtSlid.Enabled = True
                        Anato_Result.SpinButton1.Enabled = True
                        Anato_Result.frmPhoto.Enabled = True
                    
                    ElseIf optAdditional.Value = True Then
                        Anato_Result.cmdEyeCheck.Enabled = False
                '        Anato_Result.cmdGross.Enabled = False
                        Anato_Result.cmdJCode.Enabled = True
                        Anato_Result.cmdPreliminary.Enabled = False
                        Anato_Result.cmdFirstDiag.Enabled = False
                        Anato_Result.cmdSignOut.Enabled = False
                        Anato_Result.cmdAdditional.Enabled = True
                        Anato_Result.cmdJCode.Enabled = False
                        Anato_Result.txtSlid.Enabled = False
                        Anato_Result.SpinButton1.Enabled = False
                        Anato_Result.frmPhoto.Enabled = False
                    End If
    End Select
    
    
    Set GobjectSS = Anato_Jeobsu_View.ssRecept

    'Unload Me
    SSPanel1.Visible = False

    Me.Hide
    
End Sub



Private Sub optAdditional_Click(Value As Integer)
'    optJeobsu.Value = False
'    optEyeCheck.Value = False
'    optPreliminary.Value = False
'    optDiag.Value = False
'    optAdditional.Value = True
    
    SSPanel1.Visible = True
    RCheck = ""
    Call SSInitialize(ssRecept)


End Sub


Private Sub optCytology_Click(Value As Integer)
    optEyeCheck.Enabled = False
    optJeobsu.Caption = "판독입력및 수정"
    optEyeCheck.Caption = ""

End Sub

Private Sub optDiag_Click(Value As Integer)
'    optJeobsu.Value = False
'    optEyeCheck.Value = False
'    optPreliminary.Value = False
'    optDiag.Value = True
'    optAdditional.Value = False
    SSPanel1.Visible = False
    RCheck = ""
    Call SSInitialize(ssRecept)

' 스텍
End Sub

Private Sub optEyeCheck_Click(Value As Integer)
'    optJeobsu.Value = False
'    optEyeCheck.Value = True
'    optPreliminary.Value = False
'    optDiag.Value = False
'    optAdditional.Value = False
    SSPanel1.Visible = False
    RCheck = ""
    Call SSInitialize(ssRecept)

End Sub

Private Sub optHistology_Click(Value As Integer)
    optEyeCheck.Enabled = True
    optJeobsu.Caption = "육안결과입력및 수정"
    optEyeCheck.Caption = "Preliminary 입력"
    
End Sub

Private Sub optJeobsu_Click(Value As Integer)
'    optJeobsu.Value = True
'    optEyeCheck.Value = False
'    optPreliminary.Value = False
'    optDiag.Value = False
'    optAdditional.Value = False
    SSPanel1.Visible = False
    RCheck = "1"
    
    Call SSInitialize(ssRecept)


End Sub

Private Sub optPreliminary_Click(Value As Integer)
'    optJeobsu.Value = False
'    optEyeCheck.Value = False
'    optPreliminary.Value = True
'    optDiag.Value = False
'    optAdditional.Value = False
    SSPanel1.Visible = False
    RCheck = "2"
    Call SSInitialize(ssRecept)

End Sub

Private Sub optRefferal_Click(Value As Integer)
'    optEyeCheck.Enabled = True
    
    optEyeCheck.Enabled = True
    optJeobsu.Caption = "육안결과입력및 수정"
    optEyeCheck.Caption = "Preliminary 입력"
    
    
End Sub

Private Sub ssRecept_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim i                   As Integer
    
    If Row = 0 And Col = 1 Then
        ssRecept.Col = 1
        ssRecept.Row = 0
        If ssRecept.Text = "A" Then
            ssRecept.Col = 1
            ssRecept.Row = 0
            ssRecept.Text = "C"
            For i = 1 To ssRecept.DataRowCnt
                ssRecept.Row = i
                ssRecept.Text = "0"
            Next i
        Else
            ssRecept.Col = 1
            ssRecept.Row = 0
            ssRecept.Text = "A"
            For i = 1 To ssRecept.DataRowCnt
                ssRecept.Row = i
                ssRecept.Text = "1"
            Next i
        End If
    End If


End Sub

'Private Sub ssRecept_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'
'    If Button = 2 Then
'        PopupMenu mnuRecept
'    End If
'
'End Sub


Private Sub txtClass_GotFocus()

    txtClass.SelStart = 0
    txtClass.SelLength = Len(txtClass.Text)

End Sub


Private Sub TXTCLASS_KeyPress(KeyAscii As Integer)
    
    If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
    
    If KeyAscii > 0 Then
        TabCheck = TabCheck + 1
    End If
    
    If TabCheck = 1 Then
        TabCheck = 0
        SendKeys "{tab}"
    End If

End Sub


Private Sub txtDateYY_GotFocus()

    txtDateYY.SelStart = 0
    txtDateYY.SelLength = Len(txtDateYY.Text)

End Sub


Private Sub txtDATEYY_KeyPress(KeyAscii As Integer)
    
    If KeyAscii > 0 And KeyAscii <> 8 Then
        TabCheck = TabCheck + 1
    Else
        TabCheck = TabCheck - 1
    End If
    
    If TabCheck = 4 Then
        TabCheck = 0
        SendKeys "{tab}"
    End If
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    

    SendKeys "{tab}"

End Sub


Private Sub txtSeqnum_GotFocus()

    txtSeqnum.SelStart = 0
    txtSeqnum.SelLength = Len(txtSeqnum.Text)

End Sub


Private Sub txtSeqnum_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    
    SendKeys "{tab}"

End Sub





'미사용

Private Sub mnuPath_Click()
    Dim rs                  As ADODB.Recordset
    
    Dim i                   As Integer
    
    GsHistology = "YES"
    GsCytology = "NO"
    GsGross = "NO"
    GsFirst = "NO"
    GsComplete = "NO"
    GsJSHistology = "NO"
    GsJSCytology = "NO"
    
    If GsComplete = "YES" Then
        frmComplete.Enabled = True
    Else
        frmComplete.Enabled = False
    End If
    
    Call SSInitialize(ssRecept)
    
    strSQL = ""
    strSQL = strSQL & " SELECT a.*, a.RowID,"
    strSQL = strSQL & "        TO_CHAR(a.Jdate,   'YYYY-MM-DD') jdate,"
    strSQL = strSQL & "        TO_CHAR(a.Orderdt, 'YYYY-MM-DD') Orderdt"
    strSQL = strSQL & " WHERE  GBRESULT <> '9' "
    strSQL = strSQL & " AND    CLASS  = 'P' "
    strSQL = strSQL & " ORDER  BY CLASS, DATEYY, SEQNUM ASC "
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    Do Until rs.EOF
        ssRecept.Row = ssRecept.DataRowCnt + 1
        
        ssRecept.Row = i + 1
        ssRecept.Col = 2:  ssRecept.Text = rs.Fields("Class").Value & "-" & _
                                           rs.Fields("Dateyy").Value & "-" & _
                                           rs.Fields("Seqnum").Value & ""
        ssRecept.Col = 3:  ssRecept.Text = rs.Fields("Ptno").Value & ""
        ssRecept.Col = 4:  ssRecept.Text = rs.Fields("Sname").Value & ""
        
        If rs.Fields("GbGross").Value & "" = "1" Then
            ssRecept.Col = 5:    ssRecept.Text = "GROSS"
        End If
        
        Select Case rs.Fields("GbResult").Value & ""
            Case "0"
                    If rs.Fields("GbGross").Value & "" = "1" Then
                        ssRecept.Col = 5: ssRecept.Text = "GROSS"
                    Else
                        ssRecept.Col = 5: ssRecept.Text = "접수중"
                    End If
            Case "1"
                    ssRecept.Col = 5:    ssRecept.Text = "임시저장"
            Case "2"
                    ssRecept.Col = 5:    ssRecept.Text = "판독"
            Case Else
        End Select
        ssRecept.Col = 6:  ssRecept.Text = rs.Fields("OrderDt").Value & ""
        ssRecept.Col = 7:  ssRecept.Text = rs.Fields("GbGross").Value & ""
        ssRecept.Col = 8:  ssRecept.Text = rs.Fields("GbResult").Value & ""
        ssRecept.Col = 9:  ssRecept.Text = rs.Fields("RowID").Value & ""
        ssRecept.Col = 10: ssRecept.Text = rs.Fields("jdate").Value & ""
        rs.MoveNext
    Loop
    AdoCloseSet rs


End Sub

Private Sub mnuStatus_Click()

    Dim rs                  As ADODB.Recordset
    
    Dim i                   As Integer
    
    Call SSInitialize(ssRecept)
    
    strSQL = ""
    strSQL = strSQL & " SELECT a.*, a.RowID,"
    strSQL = strSQL & "        TO_CHAR(a.Jdate,   'YYYY-MM-DD') jdate,"
    strSQL = strSQL & "        TO_CHAR(a.Orderdt, 'YYYY-MM-DD') Orderdt"
    strSQL = strSQL & " WHERE  GBRESULT <> '9'           "
    strSQL = strSQL & " ORDER  BY GBRESULT, GBGROSS  ASC "
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    Do Until rs.EOF
        ssRecept.Row = ssRecept.DataRowCnt + 1
        
        ssRecept.Row = i + 1
        ssRecept.Col = 2:  ssRecept.Text = rs.Fields("Class").Value & "-" & _
                                           rs.Fields("Dateyy").Value & "-" & _
                                           rs.Fields("Seqnum").Value & ""
        ssRecept.Col = 3:  ssRecept.Text = rs.Fields("Ptno").Value & ""
        ssRecept.Col = 4:  ssRecept.Text = rs.Fields("Sname").Value & ""
        
        If rs.Fields("GbGross").Value & "" = "1" Then
            ssRecept.Col = 5:    ssRecept.Text = "GROSS"
        End If
        
        Select Case rs.Fields("GbResult").Value & ""
            Case "0"
                    If rs.Fields("GbGross").Value & "" = "1" Then
                        ssRecept.Col = 5: ssRecept.Text = "GROSS"
                    Else
                        ssRecept.Col = 5: ssRecept.Text = "접수중"
                    End If
            Case "1"
                    ssRecept.Col = 5:    ssRecept.Text = "임시저장"
            Case "2"
                    ssRecept.Col = 5:    ssRecept.Text = "판독"
            Case Else
        End Select
        ssRecept.Col = 6:  ssRecept.Text = rs.Fields("OrderDt").Value & ""
        ssRecept.Col = 7:  ssRecept.Text = rs.Fields("GbGross").Value & ""
        ssRecept.Col = 8:  ssRecept.Text = rs.Fields("GbResult").Value & ""
        ssRecept.Col = 9:  ssRecept.Text = rs.Fields("RowID").Value & ""
        ssRecept.Col = 10: ssRecept.Text = rs.Fields("jdate").Value & ""
        rs.MoveNext
    Loop
    AdoCloseSet rs
    

End Sub



Private Sub mnuTempSave_Click()
    
    Dim rs                  As ADODB.Recordset
    
    Dim i                   As Integer
    
    Call SSInitialize(ssRecept)
    
    strSQL = ""
    strSQL = strSQL & " SELECT a.*, a.RowID,"
    strSQL = strSQL & "        TO_CHAR(a.Jdate,   'YYYY-MM-DD') jdate,"
    strSQL = strSQL & "        TO_CHAR(a.Orderdt, 'YYYY-MM-DD') Orderdt"
    strSQL = strSQL & " WHERE  GBRESULT =  '1'              "
    strSQL = strSQL & " ORDER  BY CLASS, DATEYY, SEQNUM ASC "
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    Do Until rs.EOF
        ssRecept.Row = ssRecept.DataRowCnt + 1
        
        ssRecept.Row = i + 1
        ssRecept.Col = 2:  ssRecept.Text = rs.Fields("Class").Value & "-" & _
                                           rs.Fields("Dateyy").Value & "-" & _
                                           rs.Fields("Seqnum").Value & ""
        ssRecept.Col = 3:  ssRecept.Text = rs.Fields("Ptno").Value & ""
        ssRecept.Col = 4:  ssRecept.Text = rs.Fields("Sname").Value & ""
        
        If rs.Fields("GbGross").Value & "" = "1" Then
            ssRecept.Col = 5:    ssRecept.Text = "GROSS"
        End If
        
        Select Case rs.Fields("GbResult").Value & ""
            Case "0"
                    If rs.Fields("GbGross").Value & "" = "1" Then
                        ssRecept.Col = 5: ssRecept.Text = "GROSS"
                    Else
                        ssRecept.Col = 5: ssRecept.Text = "접수중"
                    End If
            Case "1"
                    ssRecept.Col = 5:    ssRecept.Text = "임시저장"
            Case "2"
                    ssRecept.Col = 5:    ssRecept.Text = "판독"
            Case Else
        End Select
        ssRecept.Col = 6:  ssRecept.Text = rs.Fields("OrderDt").Value & ""
        ssRecept.Col = 7:  ssRecept.Text = rs.Fields("GbGross").Value & ""
        ssRecept.Col = 8:  ssRecept.Text = rs.Fields("GbResult").Value & ""
        ssRecept.Col = 9:  ssRecept.Text = rs.Fields("RowID").Value & ""
        ssRecept.Col = 10: ssRecept.Text = rs.Fields("jdate").Value & ""
        rs.MoveNext
    Loop
    AdoCloseSet rs
    

End Sub



