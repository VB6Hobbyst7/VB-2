VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Anato_EyePhoto_View 
   BorderStyle     =   0  '없음
   Caption         =   "육안사진조회"
   ClientHeight    =   5940
   ClientLeft      =   -30
   ClientTop       =   2145
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5940
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '최대화
   Begin FPSpread.vaSpread ssResult 
      Height          =   7548
      Left            =   96
      TabIndex        =   0
      Top             =   756
      Width           =   10032
      _Version        =   196608
      _ExtentX        =   17695
      _ExtentY        =   13314
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
      GridColor       =   8421376
      MaxCols         =   12
      MaxRows         =   600
      ShadowColor     =   12632256
      ShadowDark      =   8421504
      ShadowText      =   0
      SpreadDesigner  =   "ANATO116.frx":0000
      VisibleCols     =   12
      VisibleRows     =   500
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   1  '위 맞춤
      Height          =   705
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12060
      _Version        =   65536
      _ExtentX        =   21272
      _ExtentY        =   1244
      _StockProps     =   15
      Caption         =   "육  안  사  진  환  자  조  회"
      ForeColor       =   8388608
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   20.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Font3D          =   2
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1785
      Left            =   10200
      TabIndex        =   2
      Top             =   780
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   3149
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
      BorderWidth     =   0
      BevelInner      =   1
      Begin MSComCtl2.DTPicker dtFromJeobsu 
         Height          =   315
         Left            =   180
         TabIndex        =   8
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   25362435
         CurrentDate     =   36311
      End
      Begin MSComCtl2.DTPicker dtToJeobsu 
         Height          =   315
         Left            =   180
         TabIndex        =   9
         Top             =   1380
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   25362435
         CurrentDate     =   36311
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
         TabIndex        =   5
         Top             =   1140
         Width           =   210
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
         TabIndex        =   4
         Top             =   510
         Width           =   420
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00808000&
         BorderStyle     =   1  '단일 고정
         Caption         =   "접수일자"
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
         TabIndex        =   3
         Top             =   120
         Width           =   1485
      End
   End
   Begin Threed.SSCommand cmdView 
      Height          =   900
      Left            =   10200
      TabIndex        =   7
      Top             =   2832
      Width           =   1692
      _Version        =   65536
      _ExtentX        =   2984
      _ExtentY        =   1587
      _StockProps     =   78
      Caption         =   "조 회"
      ForeColor       =   8388736
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      Font3D          =   3
      RoundedCorners  =   0   'False
      AutoSize        =   1
      Picture         =   "ANATO116.frx":2100
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   900
      Left            =   10200
      TabIndex        =   6
      Top             =   3792
      Width           =   1692
      _Version        =   65536
      _ExtentX        =   2984
      _ExtentY        =   1587
      _StockProps     =   78
      Caption         =   "종 료"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      Font3D          =   3
      RoundedCorners  =   0   'False
      AutoSize        =   1
      Picture         =   "ANATO116.frx":2552
   End
End
Attribute VB_Name = "Anato_EyePhoto_View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
    Unload Me

End Sub

Private Sub cmdView_Click()
    
    Dim LsPtNo              As String * 8
    Dim LsStatus            As String * 1
    Dim LsCodeKy            As String
    Dim LsDrCode            As String * 6
    Dim LsDeptCode          As String * 4
    Dim LiReccnt            As Integer
    Dim i                   As Integer
    Dim LsRet
    Dim LsDIAGNO            As String
    Dim LiPos               As Integer
    Dim LiLen               As Integer
    Dim LsChr               As String
    Dim LsStr               As String
    Dim LF
    Dim CR
    
    LF = Chr(10)
    CR = Chr(13)
    
    Call SSInitialize(ssResult)
    
    gSFrDate = Format(dtFromJeobsu.Value, "yyyy-MM-dd")
    gSToDate = Format(dtToJeobsu.Value, "yyyy-MM-dd")
    
    
    strSQL = ""
    strSQL = strSQL & " SELECT a.*,"
    strSQL = strSQL & "        TO_CHAR(a.Jdate,   'YYYY-MM-DD') Jdate1,"
    strSQL = strSQL & "        TO_CHAR(a.Orderdt, 'YYYY-MM-DD') Orderdt,"
    strSQL = strSQL & "        b.Deptnamek, c.Drname"
    strSQL = strSQL & " FROM   TWANAT_Diag  a,"
    strSQL = strSQL & "        TWBAS_Dept   b,"
    strSQL = strSQL & "        TWBAS_Doctor c "
    strSQL = strSQL & " WHERE  a.Photo     = 'Y'"
    strSQL = strSQL & " AND    a.GbResult  >= '4'"
    strSQL = strSQL & " AND    a.Jdate    BETWEEN  TO_DATE('" & gSFrDate & "','yyyy-MM-dd')"
    strSQL = strSQL & "                       AND  TO_DATE('" & gSToDate & "','yyyy-MM-dd')"
    strSQL = strSQL & " AND    a.Deptcode  = b.Deptcode(+)"
    strSQL = strSQL & " AND    a.Drcode    = c.Drcode(+)"
    strSQL = strSQL & " ORDER BY JDATE, CLASS, DATEYY, SEQNUM"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    Do Until rs.EOF
        ssResult.Row = ssResult.DataRowCnt + 1
        LiPos = 0
        LsChr = ""
        LsStr = ""
        LsDIAGNO = ""
        ssResult.Col = 1:  ssResult.Text = ssResult.Row
        
        ssResult.Col = 2:  ssResult.Text = rs.Fields("Class").Value & "-" & _
                                           rs.Fields("Dateyy").Value & "-" & _
                                           rs.Fields("Seqnum").Value & ""
        ssResult.Col = 3:  ssResult.Text = rs.Fields("Ptno").Value & ""
        ssResult.Col = 4:  ssResult.Text = rs.Fields("Sname").Value & ""
        ssResult.Col = 5:  ssResult.Text = Replace(rs.Fields("Diagno").Value & "", vbCrLf, "", 1, -1, vbTextCompare)
        
        ssResult.RowHeight(ssResult.Row) = ssResult.MaxTextRowHeight(ssResult.Row)
        
        ssResult.Col = 6:  ssResult.Text = IIf(rs.Fields("Sex").Value & "" = "M", "남", "여")
        ssResult.Col = 7:  ssResult.Text = rs.Fields("AgeYY").Value & ""
        ssResult.Col = 8:  ssResult.Text = rs.Fields("Jdate").Value & ""
        ssResult.Col = 9:  ssResult.Text = rs.Fields("OrderDt").Value & ""
        ssResult.Col = 10: ssResult.Text = rs.Fields("Roomcode").Value & ""
        ssResult.Col = 11: ssResult.Text = rs.Fields("Deptnamek").Value & ""
        ssResult.Col = 12: ssResult.Text = rs.Fields("Drname").Value & ""
        rs.MoveNext
    Loop
    ssResult.MaxRows = Rowindicator + 1
    AdoCloseSet rs
    
  
End Sub


Private Sub Form_Load()
    
    dtFromJeobsu.Value = Format(CDate(Dual_Date_Get("yyyy-MM-dd")) - 3, "yyyy-MM-dd")
    dtToJeobsu.Value = Dual_Date_Get("yyyy-MM-dd")

End Sub


