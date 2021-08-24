VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Anato_Specimen_View 
   BorderStyle     =   0  '없음
   Caption         =   "보관검체조회"
   ClientHeight    =   8715
   ClientLeft      =   1605
   ClientTop       =   1860
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8715
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '최대화
   Begin VB.Frame Frame1 
      Height          =   2805
      Left            =   9690
      TabIndex        =   8
      Top             =   2550
      Width           =   2175
      Begin Threed.SSCommand cmdDel 
         Height          =   870
         Left            =   60
         TabIndex        =   11
         Top             =   1020
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   1535
         _StockProps     =   78
         Caption         =   "삭  제"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         Font3D          =   3
         AutoSize        =   1
         MouseIcon       =   "ANATO120.frx":0000
         Picture         =   "ANATO120.frx":031A
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   870
         Left            =   60
         TabIndex        =   9
         Top             =   1890
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   1535
         _StockProps     =   78
         Caption         =   "종 료"
         ForeColor       =   255
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
         Picture         =   "ANATO120.frx":0634
      End
      Begin Threed.SSCommand cmdView 
         Height          =   870
         Left            =   60
         TabIndex        =   10
         Top             =   150
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   1535
         _StockProps     =   78
         Caption         =   "조 회"
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
         Picture         =   "ANATO120.frx":094E
      End
   End
   Begin FPSpread.vaSpread ssResult 
      Height          =   7755
      Left            =   30
      TabIndex        =   0
      Top             =   780
      Width           =   9585
      _Version        =   196608
      _ExtentX        =   16907
      _ExtentY        =   13679
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
      MaxCols         =   13
      MaxRows         =   600
      ShadowColor     =   12632256
      ShadowDark      =   8421504
      ShadowText      =   0
      SpreadDesigner  =   "ANATO120.frx":0DA0
      VisibleCols     =   12
      VisibleRows     =   500
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   1  '위 맞춤
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12060
      _Version        =   65536
      _ExtentX        =   21272
      _ExtentY        =   1296
      _StockProps     =   15
      Caption         =   "보  관  검  체  조  회"
      ForeColor       =   8388608
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   15.75
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
      Height          =   1605
      Left            =   9720
      TabIndex        =   2
      Top             =   810
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   2831
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
      Begin MSComCtl2.DTPicker dtToJeobsu 
         Height          =   315
         Left            =   660
         TabIndex        =   7
         Top             =   900
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24510467
         CurrentDate     =   36313
      End
      Begin MSComCtl2.DTPicker dtFromJeobsu 
         Height          =   315
         Left            =   660
         TabIndex        =   6
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   24510467
         CurrentDate     =   36313
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
         Left            =   150
         TabIndex        =   5
         Top             =   900
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
         Left            =   150
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
         Left            =   150
         TabIndex        =   3
         Top             =   120
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Anato_Specimen_View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdDel_Click()

    If ssResult.DataRowCnt = 0 Then Exit Sub
    
    Dim FCheck              As Boolean
    
    For i = 1 To ssResult.DataRowCnt
        ssResult.Row = i
        ssResult.Col = 1
        If ssResult.Text = "1" Then
            FCheck = True
        End If
    Next i
    
    If FCheck = False Then Exit Sub
    
    If vbNo = MsgBox("선택한 병리번호의 검체 보관을 취소합니까?", vbYesNo + vbExclamation + vbDefaultButton2, "알림") Then
        Exit Sub
    End If

'    LiDataRowCnt = ssResult.DataRowCnt
    
    For i = 1 To ssResult.DataRowCnt
        ssResult.Row = i
        ssResult.Col = 2
        If ssResult.Text = "1" Then
            ''''''''''''''''''''''''''''
            GoSub SUB_CANCEL_PROC
        End If
    Next i
    
    Exit Sub
    
'/------------------------------------------------------------------------------------------
SUB_CANCEL_PROC:
    Dim sRowID              As String
    
    ssResult.Col = 13
    sRowID = ssResult.Text
    
    strSQL = ""
    strSQL = strSQL & " UPDATE TWANAT_DIAG"
    strSQL = strSQL & " SET    SPEGEOMCH = '2'"             'Del
    strSQL = strSQL & " WHERE  RowID = '" & sRowID & "'"
    
    adoConnect.BeginTrans
    
    Result = AdoExecute(strSQL)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Result And Rowindicator > 0 Then
        
        adoConnect.CommitTrans
    Else
        adoConnect.RollbackTrans
        MsgBox " TWANAT_DIAG TABLE " & vbCrLf & _
               " Update Error가 발생하였습니다.", vbCritical, "오류"
    End If
    
    Return
    

End Sub

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
    
    gSFrDate = Format(dtFromJeobsu.Value, "yyyy-MM-dd")
    gSToDate = Format(dtToJeobsu.Value, "yyyy-MM-dd")
    
    Call SSInitialize(ssResult)
    
    strSQL = ""
    strSQL = strSQL & " SELECT a.*, a.RowID, "
    strSQL = strSQL & "        TO_CHAR(a.Jdate,   'YYYY-MM-DD') Jdate1,"
    strSQL = strSQL & "        TO_CHAR(a.DiagDate,'YYYY-MM-DD') DiagDate,"
    strSQL = strSQL & "        TO_CHAR(a.OrderDt, 'YYYY-MM-DD') OrderDt,"
    strSQL = strSQL & "        b.Deptnamek, c.Drname"
    strSQL = strSQL & " FROM   TWANAT_Diag  a,"
    strSQL = strSQL & "        TWBAS_Dept   b,"
    strSQL = strSQL & "        TWBAS_Doctor c "
    strSQL = strSQL & " WHERE  a.Spegeomch = '1'"
    strSQL = strSQL & " AND    a.GbResult  >= '4'"
    strSQL = strSQL & " AND    a.Jdate    BETWEEN  TO_DATE('" & gSFrDate & "','yyyy-MM-dd')"
    strSQL = strSQL & "                       AND  TO_DATE('" & gSToDate & "','yyyy-MM-dd')"
    strSQL = strSQL & " AND    a.Deptcode  = b.Deptcode(+)"
    strSQL = strSQL & " AND    a.Drcode    = c.Drcode(+)"
    strSQL = strSQL & " ORDER BY Jdate, CLass, DateYY, Seqnum"
    
    Result = AdoOpenSet(rs, strSQL)
    
    If Result = False Then Exit Sub
    
    Do Until rs.EOF
        ssResult.Row = ssResult.DataRowCnt + 1
        LiPos = 0
        LsChr = ""
        LsStr = ""
        LsDIAGNO = ""
        ssResult.Col = 1:  ssResult.Text = ssResult.Row
        ssResult.Col = 3:  ssResult.Text = rs.Fields("CLass").Value & "-" & _
                                           rs.Fields("Dateyy").Value & "-" & _
                                           rs.Fields("Seqnum").Value & ""
        ssResult.Col = 4:  ssResult.Text = rs.Fields("Ptno").Value & ""
        ssResult.Col = 5:  ssResult.Text = rs.Fields("Sname").Value & ""
        ssResult.Col = 6:  ssResult.Text = Replace(rs.Fields("Diagno").Value & "", vbCrLf, "", 1, -1, vbTextCompare)
        
        ssResult.RowHeight(ssResult.Row) = ssResult.MaxTextRowHeight(ssResult.Row)
        
        ssResult.Col = 7:  ssResult.Text = IIf(rs.Fields("Sex").Value = "M", "남", "여") & "/" & rs.Fields("ageYY").Value & ""
        ssResult.Col = 8:  ssResult.Text = rs.Fields("Jdate").Value & ""
        ssResult.Col = 9:  ssResult.Text = rs.Fields("OrderDt").Value & ""
        ssResult.Col = 10: ssResult.Text = rs.Fields("RoomCode").Value & ""
        ssResult.Col = 11: ssResult.Text = rs.Fields("Deptnamek").Value & ""
        ssResult.Col = 12: ssResult.Text = rs.Fields("Drname").Value & ""
        ssResult.Col = 13: ssResult.Text = rs.Fields("RowID").Value & ""
        
        rs.MoveNext
    Loop
    ssResult.MaxRows = Rowindicator + 1
    AdoCloseSet rs

End Sub

Private Sub Form_Load()
    
    dtFromJeobsu.Value = Format(CDate(Dual_Date_Get("yyyy-MM-dd")) - 7, "yyyy-MM-dd")
    dtToJeobsu.Value = Dual_Date_Get("yyyy-MM-dd")
    
End Sub

Private Sub SSCommand1_Click()




End Sub
