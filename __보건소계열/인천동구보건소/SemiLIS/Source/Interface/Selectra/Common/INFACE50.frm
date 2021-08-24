VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form INTsearch50 
   BorderStyle     =   0  '없음
   Caption         =   "검사결과 조회 및 수정"
   ClientHeight    =   7845
   ClientLeft      =   1305
   ClientTop       =   1125
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MousePointer    =   4  '아이콘
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7845
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   Begin FPSpread.vaSpread spdsearch 
      Height          =   5970
      Left            =   630
      OleObjectBlob   =   "INFACE50.frx":0000
      TabIndex        =   2
      Top             =   1155
      Width           =   10665
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   1140
      Left            =   630
      TabIndex        =   10
      Top             =   0
      Width           =   3420
      _Version        =   65536
      _ExtentX        =   6032
      _ExtentY        =   2011
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cboServer 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         Style           =   2  '드롭다운 목록
         TabIndex        =   16
         Top             =   675
         Width           =   2145
      End
      Begin VB.TextBox txtmmdd 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1050
         MaxLength       =   4
         TabIndex        =   11
         Top             =   270
         Width           =   585
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   345
         Left            =   180
         TabIndex        =   12
         Top             =   270
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   609
         _StockProps     =   15
         Caption         =   "월일입력"
         ForeColor       =   65535
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         BorderWidth     =   2
         BevelOuter      =   0
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   870
         Left            =   2475
         TabIndex        =   13
         Top             =   195
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   1535
         _StockProps     =   78
         Caption         =   "조   회"
         ForeColor       =   16576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "INFACE50.frx":047F
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   30
      Top             =   6360
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   1140
      Left            =   4410
      TabIndex        =   4
      Top             =   0
      Width           =   1860
      _Version        =   65536
      _ExtentX        =   3281
      _ExtentY        =   2011
      _StockProps     =   14
      ForeColor       =   16576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Font3D          =   4
      Begin VB.Label lbtotalcnt 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   555
         TabIndex        =   9
         Top             =   615
         Width           =   750
      End
      Begin VB.Label Label4 
         Alignment       =   2  '가운데 맞춤
         Caption         =   "TOTAL Sample"
         Height          =   195
         Left            =   210
         TabIndex        =   5
         Top             =   255
         Width           =   1500
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1140
      Left            =   6480
      TabIndex        =   3
      Top             =   0
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   2011
      _StockProps     =   14
      Caption         =   "Fast Search"
      ForeColor       =   16576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSCommand cmdUp 
         Height          =   495
         Left            =   225
         TabIndex        =   6
         Top             =   390
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   873
         _StockProps     =   78
         ForeColor       =   8421504
         BevelWidth      =   3
         Picture         =   "INFACE50.frx":1E21
      End
      Begin Threed.SSCommand cmdDown 
         Height          =   495
         Left            =   795
         TabIndex        =   7
         Top             =   390
         Width           =   555
         _Version        =   65536
         _ExtentX        =   979
         _ExtentY        =   873
         _StockProps     =   78
         BevelWidth      =   3
         Picture         =   "INFACE50.frx":213B
      End
   End
   Begin Threed.SSCommand cmdenrole 
      Height          =   870
      Left            =   8865
      TabIndex        =   1
      Top             =   150
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "등   록"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Font3D          =   2
      Picture         =   "INFACE50.frx":2455
   End
   Begin Threed.SSCommand cmdexit 
      Height          =   870
      Left            =   10440
      TabIndex        =   0
      Top             =   150
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "닫   기"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Font3D          =   2
      Picture         =   "INFACE50.frx":2A3D
   End
   Begin Threed.SSPanel pnlResult 
      Height          =   705
      Left            =   2880
      Negotiate       =   -1  'True
      TabIndex        =   8
      Top             =   5940
      Visible         =   0   'False
      Width           =   3705
      _Version        =   65536
      _ExtentX        =   6535
      _ExtentY        =   1244
      _StockProps     =   15
      Caption         =   "결과 등록 중 !!"
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Outline         =   -1  'True
      Alignment       =   2
      Begin VB.Image Imgcurrent 
         Height          =   645
         Left            =   360
         OLEDragMode     =   1  '자동
         OLEDropMode     =   2  '자동
         Stretch         =   -1  'True
         Top             =   675
         Width           =   690
      End
   End
   Begin Threed.SSCommand cmdDelete 
      Height          =   870
      Left            =   9645
      TabIndex        =   14
      Top             =   150
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "삭   제"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Font3D          =   2
      Picture         =   "INFACE50.frx":43DF
   End
   Begin Threed.SSCommand cmdResultReg 
      Height          =   645
      Left            =   870
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5970
      Width           =   2160
      _Version        =   65536
      _ExtentX        =   3810
      _ExtentY        =   1138
      _StockProps     =   78
      Caption         =   "결과  등록  준비"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   5
      Font3D          =   2
      Picture         =   "INFACE50.frx":5D81
   End
   Begin VB.Image imgbox 
      Height          =   480
      Index           =   5
      Left            =   -60
      Picture         =   "INFACE50.frx":684B
      Top             =   3600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgbox 
      Height          =   480
      Index           =   4
      Left            =   -60
      Picture         =   "INFACE50.frx":6C8D
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgbox 
      Height          =   480
      Index           =   3
      Left            =   -90
      Picture         =   "INFACE50.frx":70CF
      Top             =   2430
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgbox 
      Height          =   480
      Index           =   2
      Left            =   -90
      Picture         =   "INFACE50.frx":7511
      Top             =   1740
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgbox 
      Height          =   480
      Index           =   1
      Left            =   -90
      Picture         =   "INFACE50.frx":7953
      Top             =   1170
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgbox 
      Height          =   480
      Index           =   0
      Left            =   -60
      Picture         =   "INFACE50.frx":7D95
      Top             =   570
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "INTsearch50"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim changeKey           As Integer
Dim TestNameTable(99)   As TestNameTbl
Dim PrevSlipNo          As String
Dim scrollcnt           As Integer
Dim rotateflag          As Integer
Dim currentbox          As Integer
Dim iRow                As Integer
Dim StartBCol           As Integer
Dim EndBCol             As Integer
Dim StartBRow           As Integer
Dim EndBRow             As Integer
Dim identbOpenKey       As Integer
Dim BlockKey            As Integer
Dim CurSampCnt          As Integer
Dim CurSeqNo            As String

Private Function Append_To_Server(P_Key As String, iCnt As Integer, sOrdNo As String, RtnCd As String) As Integer
                
    Dim sLabNo  As String
    Dim II      As Integer
    Dim sqldoc  As String
    Dim sStr    As String
    Dim iRet    As Integer
    
    
'    sLabNo = Left(P_Key, 8) & Mid(P_Key, 10, 1) & Right(P_Key, 5)
    sLabNo = P_Key
        
    Append_To_Server = True
    
    '----- Server결과등록
    With Insert_Server(iCnt)
        If Trim(.ordcd) = "" Or Trim(.Result) = "" Then Exit Function
        
        If Trim(.SubNo) <> "" Then
            ReDim INSDATA(1 To 9) As String
    
            '--- Insert할 항목 조회
            sqldoc = " Select DISTINCT" _
                    & "       REQGBN, SPCGBN,  ORDDATE, DEPTCD, SEQNO" _
                    & "     , IDNO,   RSLIPCD, RORDCD,  RSPCCD " _
                    & "  from LAB_DB..LAB030M " _
                    & " where LABDATE = '" & Left(sLabNo, 8) & "'" _
                    & "   and NUMGBN = '" & Mid(sLabNo, 9, 1) & "'" _
                    & "   and LABSQNO = '" & Right(sLabNo, 5) & "'" _
                    & "   and SLIPCD = '" & Mid(.ordcd, 1, 2) & "'" _
                    & "   and ORDCD = '" & Mid(.ordcd, 3, 3) & "'" _
                    & "   and SPCCD = '" & Mid(.ordcd, 6, 2) & "'" _
                    & "   and SUBCD = ''"
'99.02.04 YEJ       & "   and ORDDATE + DEPTCD + SEQNO = '" & sOrdNo & "'" _

            If QSqlDBExec(sqldoc, QsqlConn) = QSQL_SUCCESS Then
                If QSqlGetRow(sStr, QsqlConn) = QSQL_SUCCESS Then
                    QSqlGetField 9, sStr, INSDATA()
    
                    For II = 1 To 9
                        INSDATA(II) = Trim(INSDATA(II))
                    Next II
                Else
                    iRet = QSqlSelectFree(QsqlConn)
                    'Append_To_Server = False
                    Exit Function
                End If
            Else
                iRet = QSqlSelectFree(QsqlConn)
                'Append_To_Server = False
                Exit Function
            End If
            iRet = QSqlSelectFree(QsqlConn)
            '--- 조회된 자료로 Insert처리
            sStr = " Insert into LAB_DB..LAB030M ( " _
                 & "             LABDATE, NUMGBN,  LABSQNO, SLIPCD,  ORDCD," _
                 & "             SPCCD,   SUBCD,   RSTVAL,  REFVAL,  ETCGBN," _
                 & "             REMARK,  REQGBN2, REQGBN,  SPCGBN,  ORDDATE," _
                 & "             DEPTCD,  SEQNO,   IDNO,    RSLIPCD, RORDCD," _
                 & "             RSPCCD,  RSTID,   RSTDATE) " _
                 & "    values ( '" & Left(sLabNo, 8) & "'" _
                 & "           , '" & Mid(sLabNo, 9, 1) & "'" _
                 & "           , '" & Right(sLabNo, 5) & "'" _
                 & "           , '" & Mid(.ordcd, 1, 2) & "'" _
                 & "           , '" & Mid(.ordcd, 3, 3) & "'" _
                 & "           , '" & Mid(.ordcd, 6, 2) & "'" _
                 & "           , '" & .SubNo & "'" _
                 & "           , '" & .Result & "'" _
                 & "           , '" & .Ref & "'" _
                 & "           , '', '', ''" _
                 & "           , '" & INSDATA(1) & "'" _
                 & "           , 'Y'"
            For II = 3 To 9
                sStr = sStr & ", '" & INSDATA(II) & "'"
            Next
            sStr = sStr _
                 & "           , '" & D0COM_USERID & "'" _
                 & "           , '" & Format(Now, "YYYYMMDD") & "') "
    
            If QSqlDBExec(sStr, QsqlConn) = 1 Then
                SqlStr = " Update LAB_DB..LAB030M " _
                        & "   set RSTVAL = '" & CStr(.Result) & "', " _
                        & "       REFVAL = '" & .Ref & "', " _
                        & "       RSTID  = '" & D0COM_USERID & "', " _
                        & "       RSTDATE = '" & Format(Now, "YYYYMMDD") & "'" _
                        & " where LABDATE = '" & Left(P_Key, 8) & "'" _
                        & "   and NUMGBN  = '" & Mid(P_Key, 9, 1) & "'" _
                        & "   and LABSQNO = '" & Right(P_Key, 5) & "'" _
                        & "   and SLIPCD = '" & Mid(.ordcd, 1, 2) & "'" _
                        & "   and ORDCD = '" & Mid(.ordcd, 3, 3) & "'" _
                        & "   and SPCCD = '" & Mid(.ordcd, 6, 2) & "'" _
                        & "   and SUBCD = '" & .SubNo & "'"
'99.02.04 YEJ           & "   and ORDDATE + DEPTCD + SEQNO = '" & sOrdNo & "'"
                
                If QSqlDBExec(sStr, QsqlConn) <> QSQL_SUCCESS Then
                    Append_To_Server = False
                    Exit Function
                End If
            End If
        Else
            If Mid$(P_Key, 9, 1) = "L" And .ordcd = "SL02110" And .Ref = "*" Then
                ReDim INSDATA(1 To 9) As String
                
                '--- Insert할 항목 조회
                sqldoc = " Select DISTINCT" _
                        & "       REQGBN, SPCGBN,  ORDDATE, DEPTCD, SEQNO" _
                        & "     , IDNO,   RSLIPCD, RORDCD,  RSPCCD " _
                        & "  from LAB_DB..LAB030M " _
                        & " where LABDATE = '" & Left(sLabNo, 8) & "'" _
                        & "   and NUMGBN = '" & Mid(sLabNo, 9, 1) & "'" _
                        & "   and LABSQNO = '" & Right(sLabNo, 5) & "'" _
                        & "   and SLIPCD = '" & Mid(.ordcd, 1, 2) & "'" _
                        & "   and ORDCD = '" & Mid(.ordcd, 3, 3) & "'" _
                        & "   and SPCCD = '" & Mid(.ordcd, 6, 2) & "'" _
                        & "   and SUBCD = ''"
'99.02.04 YEJ           & "   and ORDDATE + DEPTCD + SEQNO = '" & sOrdNo & "'" _

                If QSqlDBExec(sqldoc, QsqlConn) = QSQL_SUCCESS Then
                    If QSqlGetRow(sStr, QsqlConn) = QSQL_SUCCESS Then
                        QSqlGetField 9, sStr, INSDATA()
        
                        For II = 1 To 9
                            INSDATA(II) = Trim(INSDATA(II))
                        Next II
                    Else
                        iRet = QSqlSelectFree(QsqlConn)
                        'Append_To_Server = False
                        Exit Function
                    End If
                    iRet = QSqlSelectFree(QsqlConn)
                Else
                    iRet = QSqlSelectFree(QsqlConn)
                    'Append_To_Server = False
                    Exit Function
                End If
            End If
                    
            SqlStr = " Update LAB_DB..LAB030M " _
                   & "   set RSTVAL = '" & CStr(.Result) & "', " _
                   & "       REFVAL = '" & .Ref & "', " _
                   & "       RSTID  = '" & D0COM_USERID & "', " _
                   & "       RSTDATE = '" & Format(Now, "YYYYMMDD") & "'" _
                   & " where LABDATE = '" & Left(P_Key, 8) & "'" _
                   & "   and NUMGBN  = '" & Mid(P_Key, 9, 1) & "'" _
                   & "   and LABSQNO = '" & Right(P_Key, 5) & "'" _
                   & "   and SLIPCD = '" & Mid(.ordcd, 1, 2) & "'" _
                   & "   and ORDCD = '" & Mid(.ordcd, 3, 3) & "'" _
                   & "   and SPCCD = '" & Mid(.ordcd, 6, 2) & "'" _
                   & "   and SUBCD = '" & .SubNo & "'"
'99.02.04 YEJ      & "   and ORDDATE + DEPTCD + SEQNO = '" & sOrdNo & "'"
                
            If QSqlDBExec(SqlStr, QsqlConn) <> QSQL_SUCCESS Then
                Append_To_Server = False
            End If

        
            If Mid$(P_Key, 9, 1) = "L" And .ordcd = "SL02110" And .Ref = "*" Then
                '--- 조회된 자료로 Insert처리
                sStr = " Insert into LAB_DB..LAB030M ( " _
                     & "             LABDATE, NUMGBN,  LABSQNO, SLIPCD,  ORDCD," _
                     & "             SPCCD,   SUBCD,   RSTVAL,  REFVAL,  ETCGBN," _
                     & "             REMARK,  REQGBN2, REQGBN,  SPCGBN,  ORDDATE," _
                     & "             DEPTCD,  SEQNO,   IDNO,    RSLIPCD, RORDCD," _
                     & "             RSPCCD,  RSTID,   RSTDATE ) " _
                     & "    values ( '" & Left(sLabNo, 8) & "'" _
                     & "           , '" & Mid(sLabNo, 9, 1) & "'" _
                     & "           , '" & Right(sLabNo, 5) & "'" _
                     & "           , '" & "CH" & "'" _
                     & "           , '" & "021" & "'" _
                     & "           , '" & "10" & "'" _
                     & "           , '" & "" & "'" _
                     & "           , '" & "" & "'" _
                     & "           , '" & "" & "'" _
                     & "           , '', '', '' " _
                     & "           , '" & INSDATA(1) & "'" _
                     & "           , 'Y'"
                For II = 3 To 9
                    sStr = sStr _
                         & "       , '" & INSDATA(II) & "'"
                Next
                sStr = sStr _
                     & "           , '" & D0COM_USERID & "'" _
                     & "           , '') "
                Call QSqlDBExec(sStr, QsqlConn)
                
                '--- 조회된 자료로 Insert처리
                sStr = " Insert into LAB_DB..LAB030M ( " _
                     & "             LABDATE, NUMGBN,  LABSQNO, SLIPCD,  ORDCD," _
                     & "             SPCCD,   SUBCD,   RSTVAL,  REFVAL,  ETCGBN," _
                     & "             REMARK,  REQGBN2, REQGBN,  SPCGBN,  ORDDATE," _
                     & "             DEPTCD,  SEQNO,   IDNO,    RSLIPCD, RORDCD," _
                     & "             RSPCCD,  RSTID,   RSTDATE ) " _
                     & "    values ( '" & Left(sLabNo, 8) & "'" _
                     & "           , '" & Mid(sLabNo, 9, 1) & "'" _
                     & "           , '" & Right(sLabNo, 5) & "'" _
                     & "           , '" & "CH" & "'" _
                     & "           , '" & "022" & "'" _
                     & "           , '" & "10" & "'" _
                     & "           , '" & "" & "'" _
                     & "           , '" & "" & "'" _
                     & "           , '" & "" & "'" _
                     & "           , '', '', '' " _
                     & "           , '" & INSDATA(1) & "'" _
                     & "           , 'Y'"
                For II = 3 To 9
                    sStr = sStr _
                         & "       , '" & INSDATA(II) & "'"
                Next
                sStr = sStr _
                     & "           , '" & D0COM_USERID & "'" _
                     & "           , '') "
                Call QSqlDBExec(sStr, QsqlConn)
            
            ElseIf Mid$(P_Key, 9, 1) = "L" And .ordcd = "SL02110" And .Ref = "" Then
                sStr = "delete LAB_DB..LAB030M " _
                     & " where LABDATE = '" & Left(P_Key, 8) & "'" _
                     & "   and NUMGBN  = '" & Mid(P_Key, 9, 1) & "'" _
                     & "   and LABSQNO = '" & Right(P_Key, 5) & "'" _
                     & "   and SLIPCD  = 'CH' " _
                     & "   and ORDCD  in ('021', '022')" _
                     & "   and SPCCD   = '10'" _
                     & "   and SPCGBN  = 'Y'"
                Call QSqlDBExec(sStr, QsqlConn)
                
            End If
        End If
    End With

End Function

'
'   참고치 판정
'
Private Function Chk_Ref(sOrdCd As String, sSubNo As String, sRes As String, _
                        sex As String) As String

    Dim sStr    As String
    Dim sData() As String
    Dim iRet_Cd As Integer
    
    Dim LowVal  As Single
    Dim HighVal As Single
    Dim RefVal  As Single
    Dim RefChar As String

    Chk_Ref = ""

    If Not sex = "0" Then
        sStr = " Select REFLOM, REFHIM, REFCHAR, REFCHK "
    Else
        sStr = " Select REFLOF, REFHIF, REFCHAR, REFCHK "
    End If
    
    sStr = sStr & "  From BAS_DB..BAS060M " _
            & " where SLIPCD + ORDCD + SPCCD = '" & sOrdCd & "'" _
            & "   and SUBCD = '" & sSubNo & "'"
    
    If QSqlDBExec(sStr, QsqlConn) = QSQL_SUCCESS Then
        If QSqlGetRow(sStr, QsqlConn) = QSQL_SUCCESS Then
            QSqlGetField 4, sStr, sData()
            
            If Trim(sData(4)) = "C" Then    '참고치 문자
                RefChar = Trim(sData(3))
                If RefChar <> Trim(sRes) Then
                    Chk_Ref = "*"
                End If
            ElseIf Trim(sData(4)) = "N" Then        '숫자
                If sData(1) = "" And sData(2) = "" Then
                    Chk_Ref = ""
                ElseIf sData(1) = "" Then
                    RefVal = CSng(Val(Trim(sRes)))
                    HighVal = CSng(Val(sData(2)))
                
                    If RefVal > HighVal Then
                        Chk_Ref = "H"
                    End If
                ElseIf sData(2) = "" Then
                    RefVal = CSng(Val(Trim(sRes)))
                    LowVal = CSng(Val(sData(1)))
                
                    If RefVal < LowVal Then
                        Chk_Ref = "L"
                    End If
                Else
                    RefVal = CSng(Val(Trim(sRes)))
                    LowVal = CSng(Val(sData(1))): HighVal = CSng(Val(sData(2)))
                
                    If RefVal > HighVal Then
                        Chk_Ref = "H"
                    ElseIf RefVal < LowVal Then
                        Chk_Ref = "L"
                    End If
                End If
            End If
            
        End If
    End If
    iRet_Cd = QSqlSelectFree(QsqlConn)
    
End Function
Private Sub Update_DB_Result(sSamNo As String, sEqCd As String, sResult As String)

    Dim ResDB   As Database
    Dim ResTB   As Recordset
    
    If sResult <> "" Then
        Set ResDB = OpenDatabase(filename & "comm\" & strmmdd & ".mdb")
        Set ResTB = ResDB.OpenRecordset("sp_result", dbOpenTable)
        ResTB.Index = "Primarykey"
        
        ResTB.Seek "=", sSamNo, sEqCd
        If ResTB.NoMatch Then
            ResTB.AddNew
            ResTB!seq_no = sSamNo
            ResTB!testcode = sEqCd
        Else
            ResTB.Edit
        End If
        
        ResTB!TestResult = sResult
    
        ResTB.Update
        ResTB.MoveLast
        
        ResTB.Close
        ResDB.Close
    End If
End Sub

Private Sub Update_LAB020M(sOrdNo As String, sLNo As String)

    Dim iRet_Cd As Integer
    Dim sStr    As String
    Dim tData() As String
    Dim sqldoc  As String
    
    sqldoc = " Select count(*) from LAB_DB..Lab030M " _
                        & " where LABDATE = '" & Left(sLNo, 8) & "'" _
                        & "   and NUMGBN  = '" & Mid(sLNo, 9, 1) & "'" _
                        & "   and LABSQNO = '" & Right(sLNo, 5) & "'" _
                        & "   and RSTVAL = '' "
                        
    If QSqlDBExec(sqldoc, QsqlConn) = QSQL_SUCCESS Then
        If QSqlGetRow(sStr, QsqlConn) = QSQL_SUCCESS Then
        
            QSqlGetField 1, sStr, tData()
            
            If Val(tData(1)) = 0 Then
                sqldoc = " Update LAB_DB..LAB020M " _
                        & "   set ORDSTAT = '1' " _
                        & " where LABDATE = '" & Left(sLNo, 8) & "'" _
                        & "   and NUMGBN  = '" & Mid(sLNo, 9, 1) & "'" _
                        & "   and LABSQNO = '" & Right(sLNo, 5) & "'"
            Else
                sqldoc = " Update LAB_DB..LAB020M " _
                        & "   set ORDSTAT = '0' " _
                        & " where LABDATE = '" & Left(sLNo, 8) & "'" _
                        & "   and NUMGBN  = '" & Mid(sLNo, 9, 1) & "'" _
                        & "   and LABSQNO = '" & Right(sLNo, 5) & "'"
            End If
            
            iRet_Cd = QSqlDBExec(sqldoc, QsqlCode)
        End If
    End If
    iRet_Cd = QSqlSelectFree(QsqlConn)
                
End Sub
Private Sub Update_ORD020M(OrdSqNo As String)

    Dim iRet_Cd As Integer
    Dim sStr    As String
    Dim tData() As String
    Dim sqldoc  As String
    
    sqldoc = " Select count(*) from ORD_DB..ORD041M " _
            & " where ORDDATE = '" & Left(OrdSqNo, 8) & "'" _
            & "   and DEPTCD = '" & Mid(OrdSqNo, 9, 2) & "'" _
            & "   and SEQNO = '" & Right(OrdSqNo, 5) & "'" _
            & "   and RSTGBN = '' "
    
    If QSqlDBExec(sqldoc, QsqlConn) = QSQL_SUCCESS Then
        If QSqlGetRow(sStr, QsqlConn) = QSQL_SUCCESS Then
        
            QSqlGetField 1, sStr, tData()
            
            If Val(tData(1)) = 0 Then
                sqldoc = " Update ORD_DB..ORD040M " _
                        & "   set ORDSTAT = '1' " _
                        & " where ORDDATE = '" & Left(OrdSqNo, 8) & "'" _
                        & "   and DEPTCD = '" & Mid(OrdSqNo, 9, 2) & "'" _
                        & "   and SEQNO = '" & Right(OrdSqNo, 5) & "'" _
                        & "   AND ORDSTAT IN ('0','6') "
            Else
                sqldoc = " Update ORD_DB..ORD040M " _
                        & "   set ORDSTAT = '6' " _
                        & " where ORDDATE = '" & Left(OrdSqNo, 8) & "'" _
                        & "   and DEPTCD = '" & Mid(OrdSqNo, 9, 2) & "'" _
                        & "   and SEQNO = '" & Right(OrdSqNo, 5) & "'" _
                        & "   AND not ORDSTAT IN ('3','5','7') "
            
            End If
            iRet_Cd = QSqlDBExec(sqldoc, QsqlCode)
        End If
    End If
    iRet_Cd = QSqlSelectFree(QsqlConn)
                
End Sub

'
'   검사 Order 내역 Table Update
'
Private Sub Update_ORD041M(sOrdNo As String, sLNo As String)
        
    Dim kk      As Integer
    Dim iRet_Cd As Integer
    Dim sStr    As String
    Dim tData() As String
        
    For kk = 1 To 30
        With Insert_Server(kk)
            If Trim(.ordcd) <> "" And Trim(.Result) <> "" Then
            
                SqlStr = " Select count(*) from LAB_DB..Lab030M " _
                        & " where LABDATE = '" & Left(sLNo, 8) & "'" _
                        & "   and NUMGBN  = '" & Mid(sLNo, 9, 1) & "'" _
                        & "   and LABSQNO = '" & Right(sLNo, 5) & "'" _
                        & "   and RSLIPCD = '" & Mid(.RtnCd, 1, 2) & "'" _
                        & "   and RORDCD = '" & Mid(.RtnCd, 3, 3) & "'" _
                        & "   and RSPCCD = '" & Mid(.RtnCd, 6, 2) & "'" _
                        & "   and RSTVAL = '' "
                        
                If QSqlDBExec(SqlStr, QsqlConn) = QSQL_SUCCESS Then
                    If QSqlGetRow(sStr, QsqlConn) = QSQL_SUCCESS Then
                    
                        QSqlGetField 1, sStr, tData()
                        
                        If Val(tData(1)) = 0 Then
                            '--- 검사 Order 내역 Table Update
                            SqlStr = " Update ORD_DB..ORD041M " _
                                    & "   set RSTGBN = 'Y' " _
                                    & " where ORDDATE = '" & Left(sOrdNo, 8) & "'" _
                                    & "   and DEPTCD = '" & Mid(sOrdNo, 9, 2) & "'" _
                                    & "   and SEQNO = '" & Right(sOrdNo, 5) & "'" _
                                    & "   and SLIPCD = '" & Mid(.RtnCd, 1, 2) & "'" _
                                    & "   and ORDCD = '" & Mid(.RtnCd, 3, 3) & "'" _
                                    & "   and SPCCD = '" & Mid(.RtnCd, 6, 2) & "'"
                        Else
                            SqlStr = " Update ORD_DB..ORD041M " _
                                    & "   set RSTGBN = '' " _
                                    & " where ORDDATE = '" & Left(sOrdNo, 8) & "'" _
                                    & "   and DEPTCD = '" & Mid(sOrdNo, 9, 2) & "'" _
                                    & "   and SEQNO = '" & Right(sOrdNo, 5) & "'" _
                                    & "   and SLIPCD = '" & Mid(.RtnCd, 1, 2) & "'" _
                                    & "   and ORDCD = '" & Mid(.RtnCd, 3, 3) & "'" _
                                    & "   and SPCCD = '" & Mid(.RtnCd, 6, 2) & "'"
                        End If
                        iRet_Cd = QSqlDBExec(SqlStr, QsqlCode)
                    End If
                End If
                iRet_Cd = QSqlSelectFree(QsqlConn)
            End If
            
            '--- 초기화
            .ordcd = ""
            .SubNo = ""
            .Result = ""
            .Ref = ""
            .RtnCd = ""
            '-----------
        End With
    Next kk

End Sub

Private Sub Update_RegChk(sSam As String)
                
    Dim SamDb   As Database
    Dim SamTb   As Recordset
    
    Set SamDb = OpenDatabase(filename & "comm\" & strmmdd & ".mdb")
    Set SamTb = SamDb.OpenRecordset("sp_identify", dbOpenTable)
    SamTb.Index = "Primarykey"
    
    SamTb.Seek "=", sSam
    If Not SamTb.NoMatch Then
        SamTb.Edit
        SamTb!chkresult = "*"
        SamTb.Update
    End If
    
    SamTb.Close
    SamDb.Close
    
End Sub
Sub ReViewSpd()
    Dim iRet As Integer
    Screen.MousePointer = 11
    
    If identb.RecordCount > 0 Then
                                
'''''''''''''        'Spread 전체의 텍스트를 지움.
'''''''''''''            spdsearch.BlockMode = True
'''''''''''''            spdsearch.Col = 1
'''''''''''''            spdsearch.Col2 = spdsearch.MaxCols
'''''''''''''            spdsearch.Row = 1
'''''''''''''            spdsearch.Row2 = spdsearch.MaxRows
'''''''''''''            spdsearch.Action = SS_ACTION_CLEAR_TEXT
'''''''''''''            spdsearch.BlockMode = False
            
        'Spread 전체의 텍스트를 지우고 Spread의 MaxRow값을 초기화
            spdsearch.MaxRows = 0
            spdsearch.MaxRows = 20
            
            identb.Index = "primarykey"
            resulttb.Index = "Seq_No"
            
            identb.MoveFirst
    
            lbtotalcnt.Caption = Str(identb.RecordCount) & " " & "개"
            
            spdsearch.Row = 0

            '--- Index Open
            If S0SUB_Open(D0COM_SERVER01, Me.hWnd, QsqlConn) <> QSQL_SUCCESS Then GoTo DB_Close
            
'                lbtotalcnt.Caption = Str(identb.RecordCount) & " " & "개"
'                CurSampCnt = identb.RecordCount
            
            Do Until identb.EOF
                
                If Trim$(identb!slip_no & "") <> "" Then
                            '----- Server에 등록 또는 미등록 자료만 조회
                    If cboServer.ListIndex = 0 Then
                        If identb!chkresult <> "*" Or IsNull(identb!chkresult) Then       '미등록자료만 조회
                            
                            Call Row_Plus
                            Call spdsettext(spdsearch, 1, spdsearch.Row, identb!slip_no)
                            Call spdsettext(spdsearch, 6, spdsearch.Row, identb!seq_no)
                            Call Query_Data(identb!slip_no, spdsearch.Row)     'Query/Display 문
                
                            resulttb.Seek "=", identb!seq_no
                                                
                            If resulttb.NoMatch = False Then
                                Do Until resulttb.EOF
                                    If resulttb!seq_no <> identb!seq_no Then Exit Do
                                    Call spdsettext(spdsearch, TestNameTable(Val(resulttb!testcode)).col_cnt, spdsearch.Row, resulttb!TestResult)
                                    resulttb.MoveNext
                                Loop
                                
                            Else
                                Call spdsearch.SetText(5, spdsearch.Row, "0")
                            End If
                        End If
                    ElseIf cboServer.ListIndex = 1 Then
                        If identb!chkresult = "*" Then
                            Call Row_Plus
                            Call spdsettext(spdsearch, 1, spdsearch.Row, identb!slip_no)
                            Call spdsettext(spdsearch, 6, spdsearch.Row, identb!seq_no)
                            Call Query_Data(identb!slip_no, spdsearch.Row)     'Query/Display 문
                
                            resulttb.Seek "=", identb!seq_no
                                                
                            If resulttb.NoMatch = False Then
                                Do Until resulttb.EOF
                                    If resulttb!seq_no <> identb!seq_no Then Exit Do
                                    Call spdsettext(spdsearch, TestNameTable(Val(resulttb!testcode)).col_cnt, spdsearch.Row, resulttb!TestResult)
                                    resulttb.MoveNext
                                Loop
                                
                            Else
                                Call spdsearch.SetText(5, spdsearch.Row, "0")
                            End If
                            If (identb!chkresult & "") = "*" Then
                                spdsearch.Col = 1
                                spdsearch.Col2 = 1
                                spdsearch.Row = spdsearch.Row
                                spdsearch.Row2 = spdsearch.Row
                                spdsearch.BlockMode = True
                                spdsearch.BackColor = RGB(220, 220, 255)
                                spdsearch.BlockMode = False
                            End If

                        End If
                    ElseIf cboServer.ListIndex = 2 Then
                        Call Row_Plus
                        Call spdsettext(spdsearch, 1, spdsearch.Row, identb!slip_no)
                        Call spdsettext(spdsearch, 6, spdsearch.Row, identb!seq_no)
                        Call Query_Data(identb!slip_no, spdsearch.Row)     'Query/Display 문
            
                        resulttb.Seek "=", identb!seq_no
                                            
                        If resulttb.NoMatch = False Then
                            Do Until resulttb.EOF
                                If resulttb!seq_no <> identb!seq_no Then Exit Do
                                Call spdsettext(spdsearch, TestNameTable(Val(resulttb!testcode)).col_cnt, spdsearch.Row, resulttb!TestResult)
                                resulttb.MoveNext
                            Loop
                            
                        Else
                            Call spdsearch.SetText(5, spdsearch.Row, "0")
                        End If
                        If (identb!chkresult & "") = "*" Then
                            spdsearch.Col = 1
                            spdsearch.Col2 = 1
                            spdsearch.Row = spdsearch.Row
                            spdsearch.Row2 = spdsearch.Row
                            spdsearch.BlockMode = True
                            spdsearch.BackColor = RGB(220, 220, 255)
                            spdsearch.BlockMode = False
                        End If

                    End If

                End If
                
                identb.MoveNext
            
                lbtotalcnt.Caption = Str(spdsearch.Row) & " " & "개"
                CurSampCnt = identb.RecordCount
            DoEvents
            Loop
                 
            iRet = Qsqlclose(QsqlConn, ONECLOSE)
                        
            'LoadKey = False
            identbOpenKey = True
            
    Else
            
        'Spread 전체의 텍스트를 지우고 Spread의 MaxRow값을 초기화
            spdsearch.MaxRows = 0
            spdsearch.MaxRows = 20
            
            identb.Index = "primarykey"
            resulttb.Index = "Seq_No"
            
            identb.MoveFirst
    
            lbtotalcnt.Caption = Str(identb.RecordCount) & " " & "개"
            
            spdsearch.Row = 0
        
            MsgBox "해당일의 데이터가 모두 삭제되었습니다!!"
            
            
            spdsearch.EditMode = True
            spdsearch.EditMode = False
            
            txtmmdd.SetFocus
             
    End If
DB_Close:
    resulttb.Close
    identb.Close
    Db.Close
    Screen.MousePointer = 0

End Sub


Sub Row_Plus()
    If spdsearch.Row >= spdsearch.MaxRows Then
        spdsearch.MaxRows = spdsearch.MaxRows + 1
        spdsearch.Row = spdsearch.MaxRows
    Else
        spdsearch.Row = spdsearch.Row + 1
    End If
End Sub
'
'   환자 신상자료 조회 Query문
'
Private Sub Query_Data(Para As String, Row_cnt As Long)
 
    Dim tData() As String
    Dim sStr    As String
    Dim sLabNo  As String
    Dim sqldoc  As String

    With spdsearch
        '----- 환자정보 조회/표시
        sqldoc = " Select Distinct C2.PATNM, C2.SEX, D1.ORDDATE + D1.DEPTCD + D1.SEQNO, D1.RSLIPCD + D1.RORDCD + D1.RSLIPCD " _
                & "  from LAB_DB..LAB030M D1, PAT_DB..PAT010M C2 " _
                & " where D1.LABDATE = '" & Left$(Para, 8) & "'" _
                & "   and D1.NUMGBN  = '" & Mid$(Para, 10, 1) & "'" _
                & "   and D1.LABSQNO = '" & Mid$(Para, 12, 5) & "'" _
                & "   and D1.IDNO = C2.IDNO "
'                & "   and Substring(D1.ORDCD,1,2) = 'HB' " _

        If QSqlDBExec(sqldoc, QsqlConn) = QSQL_SUCCESS Then
            If QSqlGetRow(sStr, QsqlConn) = QSQL_SUCCESS Then
                        
                QSqlGetField 4, sStr, tData()
                        
                Call .SetText(2, Row_cnt, Trim(tData(1)))     '이름
                If Trim(tData(2)) = "0" Then
                    Call .SetText(3, Row_cnt, "여")     'Sex
                ElseIf Trim(tData(2)) = "1" Then
                    Call .SetText(3, Row_cnt, "남")     'Sex
                End If
                Call .SetText(4, Row_cnt, Trim(tData(3)))     'OrderNo
                Call .SetText(5, Row_cnt, Trim(tData(4)))     'RtnCode  '98.02.26 debug by KHS
                '--- 결과 표시
                'Call DIsp_Result(TableSam!SampNo, .MaxRows)
            End If
        End If
        Return_cd = QSqlSelectFree(QsqlConn)
        
    End With
    
End Sub


Private Sub cboServer_Click()
    If FrmFlag = 50 Then
        Call cmdSearch.DoClick
    End If
End Sub


Private Sub cmdDelete_Click()
    
    Dim rv As Integer
    Dim i  As Integer
    Dim j  As Integer
    Dim OldTbRows As Integer
    Dim tmpSlip
    Dim ExistTxtKey As Integer
    
    Dim CurDelCnt%
    Dim CurDelTSeq$
    Dim CurDelSeq() As String
    Dim iSseq%
    Dim iEseq%
    Dim vSeqNo
    
    If identbOpenKey = True Then
        If StartBRow = -1 And EndBRow = -1 Then
            StartBRow = 1
            EndBRow = identb.RecordCount
        End If
    End If
    
    For i = StartBRow To EndBRow
        rv = spdsearch.GetText(1, i, tmpSlip)
        If tmpSlip = "" Then
            ExistTxtKey = False
            Exit For
        Else
            ExistTxtKey = True
        End If
    Next
    
    If identbOpenKey = True And ExistTxtKey = True Then
        If StartBCol = -1 And EndBCol = -1 And BlockKey = True Then
            rv = MsgBox("블록으로 지정된 Slip을 삭제하시겠습니까?", 4, Title & "  " & "Slip No. 삭제 확인!!")
            If rv = 7 Then
                BlockKey = False
                spdsearch.EditMode = True
                spdsearch.EditMode = False
                cmdExit.SetFocus
                Exit Sub
            End If
            
            Set Db = OpenDatabase(filename & "comm\" & strmmdd & ".mdb")
            Set identb = Db.OpenRecordset("sp_identify", dbOpenTable)
            Set resulttb = Db.OpenRecordset("sp_result", dbOpenTable)

            identb.Index = "primarykey"
            resulttb.Index = "Seq_No"
            
            CurDelTSeq = ""
            identb.MoveLast
            OldTbRows = identb.RecordCount
            
            For i = Val(StartBRow) To Val(EndBRow)
                Call spdsearch.GetText(6, i, vSeqNo)
                
                identb.Seek "=", Format(CInt(vSeqNo), "0000")
                identb.Delete
                
                CurDelTSeq = CurDelTSeq & CStr(vSeqNo) & "|"
                
                resulttb.Seek "=", Format(CInt(vSeqNo), "0000")
                If resulttb.NoMatch = False Then
                   Do Until resulttb.EOF
                       If resulttb!seq_no <> Format$(CInt(vSeqNo), "0000") Then Exit Do
                       
                       resulttb.Delete
                            
                       resulttb.MoveNext
                   Loop
                End If
            Next
            
            CurDelCnt = Val(EndBRow) - Val(StartBRow) + 1
            ReDim CurDelSeq(CurDelCnt)
            
            For i = 1 To CurDelCnt
                CurDelSeq(i) = GetByOne(CurDelTSeq, CurDelTSeq)
            Next
            
            For i = 1 To CurDelCnt
                iSseq = Val(CurDelSeq(i)) + 1
                
                If i = CurDelCnt Then
                    iEseq = OldTbRows
                Else
                    iEseq = Val(CurDelSeq(i + 1)) - 1
                End If
                
                For j = iSseq To iEseq
                    identb.Seek "=", Format(j, "0000")
                    identb.Edit
                    
                    identb!seq_no = Format(j + (i * (-1)), "0000")
                    identb.Update
                    
                    resulttb.Seek "=", Format(j, "0000")
                    If resulttb.NoMatch = False Then
                       Do Until resulttb.EOF
                           If resulttb!seq_no <> Format(j, "0000") Then Exit Do

                           resulttb.Edit
                           resulttb!seq_no = Format(j + (i * (-1)), "0000")
                           resulttb.Update

                           resulttb.MoveNext
                       Loop
                    End If
                Next
            Next
'''
'''            If EndBRow <> OldTbRows Then
'''                For i = EndBRow + 1 To OldTbRows
'''                    identb.Seek "=", Format$(i, "0000")
'''                    identb.Edit
'''
'''                    identb!Seq_No = Format$(i - (EndBRow - StartBRow + 1), "0000")
'''                    identb.Update
'''
'''                    resulttb.Seek "=", Format(i, "0000")
'''                    If resulttb.NoMatch = False Then
'''                       Do Until resulttb.EOF
'''                           If resulttb!Seq_No <> Format(i, "0000") Then Exit Do
'''
'''                           resulttb.Edit
'''                           resulttb!Seq_No = Format$(i - (EndBRow - StartBRow + 1), "0000")
'''                           resulttb.Update
'''
'''                           resulttb.MoveNext
'''                       Loop
'''                    End If
'''                Next
'''            End If
        
        'Update된 테이블 새로이 보여주기
            Call ReViewSpd
            
        Else
        
            MsgBox "잘못된 삭제 방법입니다." & Chr(10) & "왼쪽의 회색빛 헤더부분을 클릭하거나 끌어서 해당줄의 전체가 어두워지게 한 후," & Chr(10) & "삭제를 하십시요!!"
        
        End If
   Else
   
        MsgBox "데이터가 없거나 조회 시작을 실행하지 않으셨습니다!!"
        
   End If
    
   BlockKey = False
   
   If identbOpenKey = False Then
        spdsearch.EditMode = True
        spdsearch.EditMode = False
   End If
   
End Sub

Private Sub cmdDown_Click()
    
    scrollcnt = spdsearch.TopRow
    If scrollcnt <= (spdsearch.MaxRows - 39) Then
        spdsearch.TopRow = scrollcnt + 20
    Else
        spdsearch.TopRow = spdsearch.MaxRows - 19
    End If
    
End Sub

Private Sub cmdexit_GotFocus()
    
    txtmmdd.SetFocus
    
End Sub

Private Sub cmdResultReg_Click()
    
    Dim tData() As String
    Dim sStr    As String
    Dim sLabNo  As String
    Dim tmpSlip
    Dim ExistTxtKey As Integer
    Dim CurRow  As Long
    Dim i       As Integer
    Dim rt      As Integer
    Dim tmpResult
    Dim sqldoc  As String
    Dim sqlConn As Long
    
   If QSqlOpen(D0COM_SERVER01, Me.hWnd, sqlConn) <> QSQL_SUCCESS Then Exit Sub
   
'    spdsearch.MaxCols = 12
    
'    Call spdsettext(spdsearch, 10, 0, "Sex")
'    Call spdsettext(spdsearch, 11, 0, "OrdNo")
'    Call spdsettext(spdsearch, 12, 0, "RtnCd")
    
'    For i = 1 To CurSampCnt
'        rt = spdsearch.GetText(1, i, tmpSlip)
'        rt = spdsearch.GetText(5, i, tmpResult)
'        If tmpSlip = "" Or tmpResult = "" Then
'            ExistTxtKey = False
''            spdsearch.MaxCols = 9
'            MsgBox "Worklist 등록 또는 검사기기로부터 결과가 전송되지 않았습니다!!"
'            Exit For
'        Else
            ExistTxtKey = True
'        End If
'    Next
    
    If ExistTxtKey = True And identbOpenKey = True Then
        
        For i = 1 To CurSampCnt
            Call spdsearch.GetText(1, i, tmpSlip)
        
            With spdsearch
                '----- 환자정보 조회/표시
                sqldoc = " Select Distinct C2.SEX, D1.ORDERNO, D1.RTNCD " _
                        & "  from LAB04_DB..DJD010M D1, LAB03_DB..DJC020M C2 " _
                        & " where D1.LABDATE = '" & Left$(tmpSlip, 8) & "'" _
                        & "   and D1.NUMGBN  = '" & Mid$(tmpSlip, 10, 1) & "'" _
                        & "   and D1.LABSQNO = '" & Mid$(tmpSlip, 12, 5) & "'" _
                        & "   and SUBSTRING(D1.ORDERNO,1,8) = C2.ORDDATE " _
                        & "   and SUBSTRING(D1.ORDERNO,9,2) = C2.DEPTCD " _
                        & "   and SUBSTRING(D1.ORDERNO,11,5) = C2.SEQNO " _
                        & "   and D1.IDNO = C2.IDNO "
                        
                If QSqlDBExec(sqldoc, sqlConn) = QSQL_SUCCESS Then
                    If QSqlGetRow(sStr, sqlConn) = QSQL_SUCCESS Then
                        
                        QSqlGetField 3, sStr, tData()
                        
                        Call .SetText(3, i, Trim(tData(1)))     'Sex
                        Call .SetText(4, i, Trim(tData(2)))     'OrderNo
                        Call .SetText(5, i, Trim(tData(3)))     'RtnCode  '98.02.26 debug by KHS
                        '--- 결과 표시
                        'Call DIsp_Result(TableSam!SampNo, .MaxRows)
                    End If
                End If
                Return_cd = QSqlSelectFree(sqlConn)
        
            End With
        
        Next
    
    End If
    
    Call Qsqlclose(sqlConn, ONECLOSE)
    
End Sub

Private Sub cmdSearch_Click()
        
    Dim i        As Integer
    Dim j        As Integer
    Dim K        As Integer
    Dim iRet     As Integer
    Dim flag_key As Integer
    Dim Row_cnt  As Long
    
    
    On Error GoTo repairDB4
    
    If IsDate(Right$(Format(Now, "yyyy"), 2) & "-" & Left$(txtmmdd, 2) & "-" & Right$(txtmmdd, 2)) = False Then
        
        MsgBox "날짜입력을 정확히 해 주세요!!"
        txtmmdd.SetFocus
    
    Else
                
        Screen.MousePointer = 11
        strmmdd = machinit & txtmmdd.Text
        textmmdd = txtmmdd.Text
        Row_cnt = 0
        identbOpenKey = False
        
        If ifFileExists(filename & "comm\" & strmmdd & ".mdb") Then
                 
             Screen.MousePointer = 11
             Set Db = OpenDatabase(filename & "comm\" & strmmdd & ".mdb")
             Set identb = Db.OpenRecordset("sp_identify", dbOpenTable)
             Set resulttb = Db.OpenRecordset("sp_result", dbOpenTable)
             
             For K = 1 To 9999
                resulttb.Index = "Seq_No"
                resulttb.Seek "=", Format(K, "0000")
                If resulttb.NoMatch = False Then
                    flag_key = True
                    Exit For
                Else
                    flag_key = False
                End If
             Next
             
             If flag_key = True Then
                                
                identb.Index = "primarykey"
                resulttb.Index = "Seq_No"
                
                identb.MoveFirst
        
                'spread 화면 초기화
                spdsearch.MaxRows = 0
                spdsearch.MaxRows = 20
                
                spdsearch.Row = 0
    
                '--- Index Open
                If S0SUB_Open(D0COM_SERVER01, Me.hWnd, QsqlConn) <> QSQL_SUCCESS Then GoTo DB_Close
                
                Do Until identb.EOF
                    
                    If Trim$(identb!slip_no & "") <> "" Then
                                '----- Server에 등록 또는 미등록 자료만 조회
                        If cboServer.ListIndex = 0 Then
                            If identb!chkresult <> "*" Or IsNull(identb!chkresult) Then       '미등록자료만 조회
                                
                                Call Row_Plus
                                Call spdsettext(spdsearch, 1, spdsearch.Row, identb!slip_no)
                                Call spdsettext(spdsearch, 6, spdsearch.Row, identb!seq_no)
                                Call Query_Data(identb!slip_no, spdsearch.Row)     'Query/Display 문
                    
                                resulttb.Seek "=", identb!seq_no
                                                    
                                If resulttb.NoMatch = False Then
                                    Do Until resulttb.EOF
                                        If resulttb!seq_no <> identb!seq_no Then Exit Do
                                        Call spdsettext(spdsearch, TestNameTable(Val(resulttb!testcode)).col_cnt, spdsearch.Row, resulttb!TestResult)
                                        resulttb.MoveNext
                                    Loop
                                    
                                Else
                                    Call spdsearch.SetText(5, spdsearch.Row, "0")
                                End If
                            End If
                        ElseIf cboServer.ListIndex = 1 Then
                            If identb!chkresult = "*" Then
                                Call Row_Plus
                                Call spdsettext(spdsearch, 1, spdsearch.Row, identb!slip_no)
                                Call spdsettext(spdsearch, 6, spdsearch.Row, identb!seq_no)
                                Call Query_Data(identb!slip_no, spdsearch.Row)     'Query/Display 문
                    
                                resulttb.Seek "=", identb!seq_no
                                                    
                                If resulttb.NoMatch = False Then
                                    Do Until resulttb.EOF
                                        If resulttb!seq_no <> identb!seq_no Then Exit Do
                                        Call spdsettext(spdsearch, TestNameTable(Val(resulttb!testcode)).col_cnt, spdsearch.Row, resulttb!TestResult)
                                        resulttb.MoveNext
                                    Loop
                                    
                                Else
                                    Call spdsearch.SetText(5, spdsearch.Row, "0")
                                End If
                                If (identb!chkresult & "") = "*" Then
                                    spdsearch.Col = 1
                                    spdsearch.Col2 = 1
                                    spdsearch.Row = spdsearch.Row
                                    spdsearch.Row2 = spdsearch.Row
                                    spdsearch.BlockMode = True
                                    spdsearch.BackColor = RGB(220, 220, 255)
                                    spdsearch.BlockMode = False
                                End If

                            End If
                        ElseIf cboServer.ListIndex = 2 Then
                            Call Row_Plus
                            Call spdsettext(spdsearch, 1, spdsearch.Row, identb!slip_no)
                            Call spdsettext(spdsearch, 6, spdsearch.Row, identb!seq_no)
                            Call Query_Data(identb!slip_no, spdsearch.Row)     'Query/Display 문
                
                            resulttb.Seek "=", identb!seq_no
                                                
                            If resulttb.NoMatch = False Then
                                Do Until resulttb.EOF
                                    If resulttb!seq_no <> identb!seq_no Then Exit Do
                                    Call spdsettext(spdsearch, TestNameTable(Val(resulttb!testcode)).col_cnt, spdsearch.Row, resulttb!TestResult)
                                    resulttb.MoveNext
                                Loop
                                
                            Else
                                Call spdsearch.SetText(5, spdsearch.Row, "0")
                            End If
                            If (identb!chkresult & "") = "*" Then
                                spdsearch.Col = 1
                                spdsearch.Col2 = 1
                                spdsearch.Row = spdsearch.Row
                                spdsearch.Row2 = spdsearch.Row
                                spdsearch.BlockMode = True
                                spdsearch.BackColor = RGB(220, 220, 255)
                                spdsearch.BlockMode = False
                            End If

                        End If

'                        Call spdsettext(spdsearch, 1, spdsearch.Row, identb!slip_no)
'                        Call spdsettext(spdsearch, 6, spdsearch.Row, identb!Seq_No)
'                        Call Query_Data(identb!slip_no, spdsearch.Row)     'Query/Display 문
'
'                        resulttb.Seek "=", identb!Seq_No
'
'                        If resulttb.NoMatch = False Then
'                            Do Until resulttb.EOF
'                                If resulttb!Seq_No <> identb!Seq_No Then Exit Do
'                                Call spdsettext(spdsearch, TestNameTable(Val(resulttb!TestCode)).col_cnt, spdsearch.Row, resulttb!TestResult)
'                                resulttb.MoveNext
'                            Loop
'
'                        Else
'                            Call spdsearch.SetText(5, spdsearch.Row, "0")
'                        End If
                
                    End If
                    
                    identb.MoveNext
                
                    lbtotalcnt.Caption = Str(spdsearch.Row) & " " & "개"
                    CurSampCnt = identb.RecordCount
'                DoEvents
                Loop
                     
                iRet = Qsqlclose(QsqlConn, ONECLOSE)
                'LoadKey = False
                identbOpenKey = True
                
             Else
                 
                 identbOpenKey = False
                 Screen.MousePointer = 0
                 
            'spread 화면 초기화
                 spdsearch.MaxRows = 0
                 spdsearch.MaxRows = 20
                 lbtotalcnt.Caption = ""
                 MsgBox "저장된 데이타가 없습니다!!"
                 Me.MousePointer = 0
                 txtmmdd.SetFocus
                 
             End If
DB_Close:
             resulttb.Close
             identb.Close
             Db.Close
        Else
                 
             Screen.MousePointer = 0
            
        'spread 화면 초기화
             spdsearch.MaxRows = 0
             spdsearch.MaxRows = 20
             lbtotalcnt.Caption = ""
             MsgBox "화일 " & strmmdd & ".mdb가 존재하지 않습니다."
             
             Me.MousePointer = 0
             txtmmdd.SetFocus
             
        End If
        
        Me.MousePointer = 0
        Screen.MousePointer = 0
                    
    End If
    
    
        
GoTo END_PROC
repairDB4:
    If Err = 3049 Then
        MsgBox "데이타가 손상되어 있습니다. 확인을 누르시면 데이타를 복구합니다."
        RepairDatabase (filename & "comm\" & strmmdd & ".mdb")
        Set dbrp = OpenDatabase(filename & "comm\" & strmmdd & ".mdb")
    End If
    Resume Next

END_PROC:
End Sub

Private Sub cmdUp_Click()
    
    scrollcnt = spdsearch.TopRow

    If scrollcnt > 20 Then
            spdsearch.TopRow = scrollcnt - 20
    End If
    
    spdsearch.TopRow = 1
End Sub

Private Sub cmdenrole_Click()
    
    Dim i%, rt%, seqnoVar, slipnoVar, tcode$, tresult$
    Dim tmpSlip
    Dim tmpResult
    Dim ExistTxtKey As Integer
    
    Dim filename    As String
    Dim iRet    As Integer

    Dim iR  As Integer
    Dim iC  As Integer
    Dim Tmp As Variant
    Dim ix1 As Integer
        
    Dim labno   As String
    Dim SampNo  As String
    Dim sSex    As String
    Dim sOrderNo    As String
    Dim sRtnCd  As String
    
    Dim ChkTrans    As Integer
    Dim ChkExist    As Integer

'    Call cmdResultReg_Click
    
'    If Val(lbtotalcnt.Caption) < 1 Then
'        Exit Sub
'    End If
        '--- Index Open
    If S0SUB_Open(D0COM_SERVER01, Me.hWnd, QsqlConn) <> QSQL_SUCCESS Then Exit Sub
    If S0SUB_Open(D0COM_SERVER01, Me.hWnd, QsqlCode) <> QSQL_SUCCESS Then Exit Sub

    MousePointer = 11
    Timer1.Enabled = True
    
    ChkExist = False
    
    For iR = 1 To Val(lbtotalcnt.Caption)
    
        spdsearch.Col = 1
        spdsearch.Row = iR
        If spdsearch.BackColor <> RGB(220, 220, 255) Then
'        If iR = 17 Then
            '서버로 등록
            '--- Check Box
            Call spdsearch.GetText(1, iR, Tmp)
            If Trim(Tmp) <> "" Then
                With spdsearch
                    Call .GetText(1, iR, Tmp)
                    labno = Left(Tmp, 8) & Mid(Tmp, 10, 1) & Mid(Tmp, 12, 5)      '접수번호
                    Call .GetText(3, iR, Tmp):
                    If Trim(Tmp) = "여" Then
                        sSex = "0"
                    ElseIf Trim(Tmp) = "남" Then
                        sSex = "1"
                    End If
                    Call .GetText(4, iR, Tmp): sOrderNo = Trim(Tmp)
                    Call .GetText(5, iR, Tmp): sRtnCd = Trim(Tmp) '1/12 YK
                    Call .GetText(6, iR, Tmp): SampNo = Trim(Tmp)          'Sample No
                End With
                If sRtnCd <> "0" Then
                    '접수번호 등록중입니다.
                    ChkExist = True
                    
                    For iC = 1 To 90
                        If Trim(TestNameTable(iC).code) <> "" Then
                            Call spdsearch.GetText(TestNameTable(iC).col_cnt, iR, Tmp)
                            
                            With Insert_Server(iC)
                                .ordcd = Left(TestNameTable(iC).code, 7)
                                If Len(Trim(TestNameTable(iC).code)) > 8 Then
                                    .SubNo = Right(Trim(TestNameTable(iC).code), 2)
                                Else
                                    .SubNo = ""
                                End If
                                '--- 결과 앞의 '$'표시 제거
                                If Left(Trim(Tmp), 1) = "$" Then
                                    .Result = Mid(Trim(Tmp), 2)
                                Else
                                    .Result = Trim(Tmp)
                                End If
                                '--------------------------
                                .RtnCd = Get_RtnCd(labno, .ordcd, .SubNo, sOrderNo)    '1/12 yk
                                
                                .Ref = Chk_Ref(.ordcd, .SubNo, .Result, sSex)
                                '--- Hi_Result 내용 Update
                                Call Update_DB_Result(SampNo, Format(iC, "00"), .Result)
                            End With
                        End If
                    Next iC
                                        
                    '----- 구조체에 저장된 결과 Server에 등록
                    ret = QSqlBeginTrans()
                    DBEngine.Workspaces(0).BeginTrans
                    ChkTrans = False
                    
                    For ix1 = 1 To 90
                        '----- 검사항목별 결과입력(Batch)
                        If Append_To_Server(labno, ix1, sOrderNo, sRtnCd) = True Then
                            ChkTrans = True
'                            Exit For
                        End If
                    Next ix1
                    
                    If ChkTrans = False Then
                        DBEngine.Workspaces(0).Rollback
                        ret = QSqlRollBack()      'TRANSACTION 에러종료
                        
                    Else
                        DBEngine.Workspaces(0).CommitTrans
                        ret = QSqlCommitTrans()    'TRANSACTION 정상종료
                        '--- 검사 Order내역 Table Update
                        Call Update_ORD041M(sOrderNo, labno)
                        
                        Call Update_LAB020M(sOrderNo, labno)
                        
                        '--- 진료과별 처방내역 Update
                        Call Update_ORD020M(sOrderNo)
                        
                        '--- 등록체크 Update(MDB)
                        Call Update_RegChk(SampNo)
                        '--- 등록된후 색 변화
                        With spdsearch
                            .Row = iR: .Row2 = iR
                            .Col = 1: .Col2 = .MaxCols
                            .BlockMode = True
                            .BackColor = RGB(220, 220, 255)
                            .BlockMode = False
                        End With
                    End If
                End If
            End If
        End If
    Next
            
    '--- Index Close
    ret = Qsqlclose(QsqlConn, ONECLOSE)
    ret = Qsqlclose(QsqlCode, ONECLOSE)

    If ChkExist <> True Then
        MsgBox "등록할 자료가 없습니다. 확인하십시오.", vbInformation
    End If
    
    Timer1.Enabled = False
    
    MousePointer = 0

End Sub

Private Sub cmdExit_Click()

    Unload Me
    FrmFlag = 0
End Sub

Private Sub Form_Activate()
    
    txtmmdd.SetFocus
    
End Sub

Private Sub Form_Load()
    
    Dim tablerows   As Integer
    Dim sRow        As Integer
    Dim TestItemNo  As Integer
    Dim i           As Integer
    
    
    'form을 위치
    Me.Top = 0
    Me.Left = 0
    Me.Height = INTmain00.Height - INTmain00.pnlMain.Height - 500
    Me.Width = INTmain00.Width - 200
    
    pnlResult.ZOrder 0    '결과등록 중 임을 나타내는 PANEL 사용키 위해
    currentbox = 0
    rotateflag = 0
    Timer1.Enabled = False
    
    '--- Combo Box Display
    With cboServer
        .AddItem ("미등록 결과자료")
        .AddItem ("등  록 결과자료")
        .AddItem ("전  체 결과자료")
        
        '조회구분별 환자 신상자료 조회/화면 표시
        .ListIndex = 0
    End With
    
    Call spdsettext(spdsearch, 3, 0, "Sex")
    Call spdsettext(spdsearch, 4, 0, "OrdNo")
    Call spdsettext(spdsearch, 5, 0, "RtnCd")
    Call spdsettext(spdsearch, 6, 0, "SampleNo")
    
    txtmmdd.Text = Format(month(Now), "00") & Format(day(Now), "00")
           
    Set dbcode = OpenDatabase(filename & codestr)
    Set tbcode = dbcode.OpenRecordset("cdtable")

    tbcode.MoveFirst

    For sRow = 1 To 90
'        TestNameTable(sRow).Name = tbcode!Name
'        TestNameTable(sRow).code = tbcode!code
'        If TestNameTable(sRow).code <> "" Then
'            TestItemNo = TestItemNo + 1
'            TestNameTable(sRow).col_cnt = TestItemNo + 6
'            spdsearch.MaxCols = TestNameTable(sRow).col_cnt
'            Call spdsettext(spdsearch, TestNameTable(sRow).col_cnt, 0, TestNameTable(sRow).Name)
'        End If
        TestNameTable(Val(tbcode!EQIPNO)).Name = tbcode!Name
        TestNameTable(Val(tbcode!EQIPNO)).code = tbcode!code
        If TestNameTable(Val(tbcode!EQIPNO)).code <> "" Then
            TestItemNo = TestItemNo + 1
            TestNameTable(Val(tbcode!EQIPNO)).col_cnt = TestItemNo + 6
            spdsearch.MaxCols = TestNameTable(Val(tbcode!EQIPNO)).col_cnt
            Call spdsettext(spdsearch, TestNameTable(Val(tbcode!EQIPNO)).col_cnt, 0, TestNameTable(Val(tbcode!EQIPNO)).Name)
        End If
        
        tbcode.MoveNext
    Next
    
   
    tbcode.Close
    dbcode.Close
    
'    If TestItemNo >= 17 Then
'        spdsearch.MaxCols = TestItemNo + 1
'        For i = 1 To TestItemNo
'            Call spdsettext(spdsearch, i + 1, 0, TestNameTable(i).name)
'        Next
'    Else
'        spdsearch.MaxCols = 9
'        For i = 1 To TestItemNo
'            Call spdsettext(spdsearch, i + 1, 0, TestNameTable(i).name)
'        Next
'
'        For i = TestItemNo + 1 To spdsearch.MaxCols - 1
'            Call spdsettext(spdsearch, i + 1, 0, "-")
'        Next
'    End If
    
'1st Column(SlipNo)의 색깔을 노란색
    spdsearch.BlockMode = True
    spdsearch.Col = 1
    spdsearch.Col2 = 1
    spdsearch.Row = -1
    spdsearch.Row2 = -1
    spdsearch.BackColor = &HC0FFFF
    spdsearch.BlockMode = False
    
    identbOpenKey = False       '아직 DB가 열리지 않았음
    FrmFlag = 50
            
    Call cmdSearch.DoClick
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim Index%
'    Index = 7
'    Call MainTitle_Bold(Index)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If identbOpenKey = True Then
'        identb.Close
'        resulttb.Close
'        db.Close
    End If
    
    'LoadKey = True
    
End Sub

Private Sub spdsearch_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    StartBCol = CInt(BlockCol)
    StartBRow = CInt(BlockRow)
    EndBCol = CInt(BlockCol2)
    EndBRow = CInt(BlockRow2)
    BlockKey = True
      
End Sub

Private Sub spdsearch_Change(ByVal Col As Long, ByVal Row As Long)
    
'    Dim RtnCd   As Boolean
'    Dim rv      As Integer
'    Dim ccode   As String
'    Dim slipnoVar
'    Dim rresult
'    Dim SEQNO
'    Dim irow    As Integer
'
'    RtnCd = spdsearch.GetText(1, Row, slipnoVar)
'    slipnoVar = Format(slipnoVar, DigitShape)
'    irow = CInt(Row)
'
'    If Row <= identb.RecordCount Then
'        If Col = 1 Then
'            If Len(slipnoVar) > SlipDigit Then
'                MsgBox "입력하신 Slip No.가" & " " & SlipDigit & "자리를 넘습니다!!"
'                Call spdsettext(spdsearch, 1, irow, PrevSlipNo)
'                Exit Sub
'            End If
'
'            If Len(slipnoVar) = 0 Then
'                MsgBox "Slip No.를 0자리로 할 수 없습니다!!" & Chr(10) & "왼쪽의 회색빛 헤더부분을 클릭하거나 끌어서 해당줄의 전체가 어두워지게 한 후," & Chr(10) & "삭제를 하십시요!!"
'                Call spdsettext(spdsearch, 1, irow, PrevSlipNo)
'                Exit Sub
'            End If
'
'            Call spdsettext(spdsearch, 1, irow, slipnoVar)
'
'            If PrevSlipNo <> slipnoVar Then
'
'                If IsNumeric(slipnoVar) = False Then
'                    rv = MsgBox("입력하신 Slip No.에 문자가 포함되어 있습니다." & Chr(10) & "문자가 포함된 Slip No.로 바꾸시겠습니까?", 4, Title & " Slip No. 변경 확인!!")
'                    If rv = 7 Then
'                        Call spdsettext(spdsearch, 1, irow, PrevSlipNo)
'                        Exit Sub
'                    End If
'                End If
'
'        'SLIP NO. 등록여부
'                rv = MsgBox("Slip No.를 변경하시겠습니까?", 4, Title & " Slip No. 변경 확인!!")
'                If rv = 7 Then
'                    Call spdsettext(spdsearch, 1, irow, PrevSlipNo)
'                    Exit Sub
'                End If
'
'                identb.Index = "primarykey"
'                identb.Seek "=", Format(Row, "0000")
'                If identb.NoMatch = False Then
'                    identb.Edit
'                    identb!slip_no = slipnoVar
'                    identb.Update
'                Else
'                    identb.AddNew
'                    identb!Seq_No = Format(Row, "0000")
'                    identb!slip_no = slipnoVar
'                    identb!ChkResult = "&"
'                    identb.Update
'                End If
'
'                If identb!ChkResult = "*" Then
'                    identb.Edit
'                    identb!ChkResult = "&"
'                    identb.Update
'                    spdsearch.BlockMode = True
'                    spdsearch.Col = 1
'                    spdsearch.col2 = 1
'                    spdsearch.Row = Row
'                    spdsearch.row2 = Row
'                    spdsearch.BackColor = &HC0FFFF
'                    spdsearch.BlockMode = False
'                End If
'
'            End If
'
'        Else
'
'            RtnCd = spdsearch.GetText(Col, Row, rresult)
'
'            ccode = Format(Col - 1, "00")
'
'            identb.Index = "primarykey"
'            identb.Seek "=", Format(Row, "0000")
'            SEQNO = identb!Seq_No
'
'            resulttb.Index = "primarykey"
'            resulttb.Seek "=", SEQNO, ccode
'
'            If rresult = "" Then
'                If resulttb.NoMatch = False Then
'                    resulttb.Delete
'                End If
'            Else
'                If resulttb.NoMatch = False Then
'                    resulttb.Edit
'                    resulttb!TestResult = rresult
'                    resulttb.Update
'                Else
'                    resulttb.AddNew
'                    resulttb!Seq_No = SEQNO
'                    resulttb!TestCode = ccode
'                    resulttb!TestResult = rresult
'                    resulttb.Update
'                End If
'            End If
'
'            If identb!ChkResult = "*" Then
'                identb.Edit
'                identb!ChkResult = "&"
'                identb.Update
'                spdsearch.BlockMode = True
'                spdsearch.Col = 1
'                spdsearch.col2 = 1
'                spdsearch.Row = Row
'                spdsearch.row2 = Row
'                spdsearch.BackColor = &HC0FFFF
'                spdsearch.BlockMode = False
'            End If
'
'        End If
'
'    End If

End Sub

Private Sub spdsearch_Click(ByVal Col As Long, ByVal Row As Long)
   
''    spdsearch.Row = spdsearch.ActiveRow
''    spdsearch.Col = spdsearch.ActiveCol
    
    spdsearch.SelStart = 0
    spdsearch.SelLength = Len(spdsearch.Text)
        
    Dim slipnoVar
    Dim fv  As Boolean
    
    'general에 선언 PrevSlipNo
    If Col = 1 Then
        fv = spdsearch.GetText(1, Row, slipnoVar)
        PrevSlipNo = slipnoVar
        
    End If

End Sub

Private Sub spdsearch_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim slipnoVar
    Dim fv  As Boolean
    
    'general에 선언 PrevSlipNo
    If Col = 1 Then
        fv = spdsearch.GetText(1, Row, slipnoVar)
        If fv = True Then
            spdsearch.Col = 1
            spdsearch.Row = Row
            If spdsearch.BackColor = RGB(220, 220, 255) Then
                spdsearch.Col = 1
                spdsearch.Col2 = 1
                spdsearch.Row = Row
                spdsearch.Row2 = Row
                spdsearch.BlockMode = True
                spdsearch.BackColor = RGB(220, 255, 220)
                spdsearch.BlockMode = False
'                PrevSlipNo = slipnoVar
            ElseIf spdsearch.BackColor = RGB(220, 255, 220) Then
                spdsearch.Col = 1
                spdsearch.Col2 = 1
                spdsearch.Row = Row
                spdsearch.Row2 = Row
                spdsearch.BlockMode = True
                spdsearch.BackColor = RGB(220, 220, 255)
                spdsearch.BlockMode = False
            End If
        End If
    End If
End Sub

Private Sub spdsearch_GotFocus()
    
    'spdsearch.Row = spdsearch.ActiveRow
    'spdsearch.Col = spdsearch.ActiveCol

    spdsearch.SelStart = 0
    spdsearch.SelLength = Len(spdsearch.Text)
        
    Dim slipnoVar
    Dim fv  As Boolean
    
    'general에 선언 PrevSlipNo
    If spdsearch.Col = 1 Then
        fv = spdsearch.GetText(1, spdsearch.Row, slipnoVar)
        PrevSlipNo = slipnoVar
        
    End If

End Sub

Private Sub Timer1_Timer()
    Dim i%
    
    i = 1
    
    If rotateflag = 1 Then
        Imgcurrent.Picture = imgbox(currentbox).Picture
        Imgcurrent.Left = 360 + currentbox * 700
        Imgcurrent.Top = 625 + (-1) ^ currentbox * 50
        If Imgcurrent.Left = 3330 Then
            Imgcurrent.Left = 360
            Imgcurrent.Top = 625
        End If
        i = i + 1
        currentbox = currentbox + 1
        If currentbox = 5 Then
            currentbox = 0
        End If
    End If
 
End Sub


Private Sub txtmmdd_Click()
    Call txbox_highlight(txtmmdd)
End Sub

Private Sub txtmmdd_GotFocus()
    Call txbox_highlight(txtmmdd)
End Sub


Private Sub txtmmdd_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim j As Integer
    
   If KeyCode = 13 Then
            On Error GoTo repairDB4
            cmdSearch_Click
'            If IsDate(Right$(Format(Now, "yyyy"), 2) & "-" & Left$(txtmmdd, 2) & "-" & Right$(txtmmdd, 2)) = False Then
'
'                MsgBox "날짜입력을 정확히 해 주세요!!"
'                txtmmdd.SetFocus
'
'            Else
'
'                Screen.MousePointer = 11
'                strmmdd = machinit & txtmmdd.Text
'                textmmdd = txtmmdd.Text
'
'                identbOpenKey = False
'
'                If ifFileExists(FileName & "comm\" & strmmdd & ".mdb") Then
'
'                     Screen.MousePointer = 11
'                     Set db = OpenDatabase(FileName & "comm\" & strmmdd & ".mdb")
'                     Set identb = db.OpenRecordset("sp_identify", dbOpenTable)
'                     Set resulttb = db.OpenRecordset("sp_result", dbOpenTable)
'
'                     If identb.RecordCount > 0 Then
'
'                        identb.Index = "primarykey"
'                        resulttb.Index = "Seq_No"
'
'                        identb.MoveFirst
'
'                        lbtotalcnt.Caption = Str(identb.RecordCount) & " " & "개"
'                        'Label5.Caption = Val(Left$(textmmdd, 2)) & "월" & " " & Val(Right$(textmmdd, 2)) & "일"
'
'                    'Spread 초기화
'                        spdsearch.MaxRows = 0
'                        spdsearch.MaxRows = 20
'
'                        spdsearch.Row = 0
'
'                        Do Until identb.EOF
'
'                            Call Row_Plus
'
'                            If Trim$(identb!slip_no & "") <> "" Then
'                                Call spdsettext(spdsearch, 1, spdsearch.Row, identb!slip_no)
'                            Else
'                                MsgBox "SlipNo가 존재하지 않습니다!!"
'                                Exit Sub
'                            End If
'
'                            resulttb.Seek "=", identb!Seq_No
'
'                            If resulttb.NoMatch = False Then
'                                Do Until resulttb.EOF
'                                    If resulttb!Seq_No <> identb!Seq_No Then Exit Do
'                                    spdsearch.Col = Val(resulttb!TestCode) + 1
'                                    Call spdsettext(spdsearch, spdsearch.Col, spdsearch.Row, resulttb!TestResult)
'                                    resulttb.MoveNext
'                                    spdsearch.Col = spdsearch.Col + 1
'                                Loop
'
'                            End If
'
'                            If (identb!ChkResult & "") = "*" Then
'                                spdsearch.Col = 1
'                                spdsearch.col2 = 1
'                                spdsearch.Row = spdsearch.Row
'                                spdsearch.row2 = spdsearch.Row
'                                spdsearch.BlockMode = True
'                                spdsearch.BackColor = RGB(220, 220, 255)
'                                spdsearch.BlockMode = False
'                            End If
'
'                            identb.MoveNext
'
'                        Loop
'
'                        'LoadKey = False
'                        identbOpenKey = True
'
'                     Else
'
'                         resulttb.Close
'                         identb.Close
'                         db.Close
'                         identbOpenKey = False
'                         Screen.MousePointer = 0
'                         MsgBox "저장된 데이타가 없습니다!!"
'                         Me.MousePointer = 0
'                         txtmmdd.SetFocus
'
'                     End If
'
'                Else
'
'
'                     Screen.MousePointer = 0
'
'                'spread 화면 초기화
'                     spdsearch.MaxRows = 0
'                     spdsearch.MaxRows = 20
'
'                     MsgBox "화일 " & strmmdd & ".mdb가 존재하지 않습니다."
'
'                     Me.MousePointer = 0
'                     'txtmmdd.SetFocus
'
'                End If
'
'                Me.MousePointer = 0
'                Screen.MousePointer = 0
'
'            End If
'
'        KeyCode = 0
'        cmdexit.SetFocus
           
    End If
    
            
GoTo END_PROC
repairDB4:
    If Err = 3049 Then
        MsgBox "데이타가 손상되어 있습니다. 확인을 누르시면 데이타를 복구합니다."
        RepairDatabase (filename & "comm\" & strmmdd & ".mdb")
        Set dbrp = OpenDatabase(filename & "comm\" & strmmdd & ".mdb")
    End If
    Resume Next

END_PROC:
End Sub


