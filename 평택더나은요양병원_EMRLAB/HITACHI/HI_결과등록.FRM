VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm결과등록 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "결과 조회 및 등록"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11955
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   11955
   ShowInTaskbar   =   0   'False
   Begin FPSpread.vaSpread spdList 
      Height          =   5445
      Left            =   120
      OleObjectBlob   =   "Hi_결과등록.frx":0000
      TabIndex        =   3
      Top             =   810
      Width           =   11565
   End
   Begin Threed.SSPanel pnlStatus 
      Height          =   495
      Left            =   330
      TabIndex        =   14
      Top             =   6420
      Width           =   4725
      _Version        =   65536
      _ExtentX        =   8334
      _ExtentY        =   873
      _StockProps     =   15
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      Font3D          =   1
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   645
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   5715
      _Version        =   65536
      _ExtentX        =   10081
      _ExtentY        =   1138
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
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   2145
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   375
         Left            =   150
         TabIndex        =   8
         Top             =   180
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   15
         Caption         =   "월일입력"
         ForeColor       =   65535
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   2
         BorderWidth     =   2
         BevelOuter      =   0
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSCommand cmdQuery 
         Height          =   435
         Left            =   3960
         TabIndex        =   6
         Top             =   150
         Width           =   1605
         _Version        =   65536
         _ExtentX        =   2831
         _ExtentY        =   767
         _StockProps     =   78
         Caption         =   "조 회     시 작"
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
         RoundedCorners  =   0   'False
         Picture         =   "Hi_결과등록.frx":18C3
      End
      Begin MSMask.MaskEdBox mskDate 
         Height          =   375
         Left            =   1020
         TabIndex        =   0
         Top             =   180
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####"
         PromptChar      =   "_"
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   645
      Left            =   8280
      TabIndex        =   9
      Top             =   60
      Width           =   3405
      _Version        =   65536
      _ExtentX        =   6006
      _ExtentY        =   1138
      _StockProps     =   14
      Caption         =   "Fast Search"
      ForeColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   1
      Begin VB.TextBox txtLabID 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         MaxLength       =   15
         TabIndex        =   2
         Top             =   210
         Width           =   1950
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "접수번호"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   285
         TabIndex        =   10
         Top             =   300
         Width           =   825
      End
   End
   Begin Threed.SSCommand cmdAppend 
      Height          =   645
      Left            =   5400
      TabIndex        =   4
      Top             =   6330
      Width           =   3015
      _Version        =   65536
      _ExtentX        =   5318
      _ExtentY        =   1138
      _StockProps     =   78
      Caption         =   "결 과             등 록"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   4
      Font3D          =   2
      RoundedCorners  =   0   'False
      Picture         =   "Hi_결과등록.frx":2715
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   645
      Left            =   8400
      TabIndex        =   5
      Top             =   6330
      Width           =   3015
      _Version        =   65536
      _ExtentX        =   5318
      _ExtentY        =   1138
      _StockProps     =   78
      Caption         =   "닫                기"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   4
      Font3D          =   2
      RoundedCorners  =   0   'False
      Picture         =   "Hi_결과등록.frx":31DF
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   645
      Left            =   5730
      TabIndex        =   11
      Top             =   60
      Width           =   2595
      _Version        =   65536
      _ExtentX        =   4577
      _ExtentY        =   1138
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
      Begin Threed.SSPanel pnlCnt 
         Height          =   375
         Left            =   1770
         TabIndex        =   13
         Top             =   180
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   661
         _StockProps     =   15
         BackColor       =   12648384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "TOTAL Sample: "
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   12
         Top             =   285
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frm결과등록"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   SampNo가 Key인 장비용 결과조회/입력 화면
'   (Hitachi 용)
'
Option Explicit

    Dim ChkTrans    As Integer
    Dim TablePtr    As Integer
    Dim R_Row   As Integer
        
    Dim ResDB_Name  As String       'Result DB명
    Dim SamDB_Name  As String       'Sample DB명

    Dim OldRow  As Integer

Private Function Append_To_Server(P_Key As String, iCnt As Integer, sOrdNo As String) As Integer
                
    Dim iRet    As Integer
    Dim sStr    As String
    Dim rData() As String
    Dim sLabNo    As String
    Dim ii      As Integer
        
    sLabNo = Left(P_Key, 8) & Mid(P_Key, 10, 1) & Right(P_Key, 5)
        
    Append_To_Server = True
    
    '----- Server결과등록
    With Insert_Server(iCnt)
        '--- 검사 Order 내역 Table Update
        SqlStr = " Update LAB03_DB..DJC050M " _
                & "   set RSTGBN = 'Y' " _
                & " where ORDDATE = '" & Left(sOrdNo, 8) & "'" _
                & "   and DEPTCD = '" & Mid(sOrdNo, 9, 2) & "'" _
                & "   and SEQNO = '" & Right(sOrdNo, 5) & "'" _
                & "   and ORDCD = '" & .OrdCd & "'"
        
        If QSqlDBExec(SqlStr, QsqlConn) = QSQL_SUCCESS Then
            '--- Update
            SqlStr = " Update LAB04_DB..DJD010M " _
                    & "   set RSTVAL = '" & .Result & "', " _
                    & "       REFVAL = '" & .Ref & "', " _
                    & "       RSTID  = '" & D0COM_USERID & "', " _
                    & "       RSTDATE = '" & Format(Now, "YYYYMMDD") & "'" _
                    & " where LABDATE = '" & Left(sLabNo, 8) & "'" _
                    & "   and NUMGBN  = '" & Mid(sLabNo, 9, 1) & "'" _
                    & "   and LABSQNO = '" & Right(sLabNo, 5) & "'" _
                    & "   and ORDCD = '" & .OrdCd & "'" _
                    & "   and SUBCD = '" & .SubNo & "'" _
                    & "   and ORDERNO = '" & sOrdNo & "'"
            
            If QSqlDBExec(SqlStr, QsqlConn) <> QSQL_SUCCESS Then
                '--- Insert(Sub 검사항목인 경우-조회 후 입력처리)
                ReDim INSDATA(1 To 5) As String
                
                '--- Insert할 항목 조회
                SqlDoc = " Select DISTINCT REQGBN, SPCGBN, RETGBN, RTNCD, IDNO " _
                        & "  from LAB04_DB..DJD010M " _
                        & " where LABDATE = '" & Left(sLabNo, 8) & "'" _
                        & "   and NUMGBN = '" & Mid(sLabNo, 9, 1) & "'" _
                        & "   and LABSQNO = '" & Right(sLabNo, 5) & "'" _
                        & "   and ORDCD = '" & .OrdCd & "'" _
                        & "   and ORDERNO = '" & sOrdNo & "'"
                        
                If QSqlDBExec(SqlDoc, QsqlConn) = QSQL_SUCCESS Then
                    If QSqlGetRow(sStr, QsqlConn) = QSQL_SUCCESS Then
                        QSqlGetField 5, sStr, rData()
                        
                        For ii = 1 To 5
                            INSDATA(ii) = Trim(rData(ii))
                        Next ii
                    Else
                        iRet = QSqlSelectFree(QsqlConn)
                        Append_To_Server = False
                        Exit Function
                    End If
                Else
                    iRet = QSqlSelectFree(QsqlConn)
                    Append_To_Server = False
                    Exit Function
                End If
                iRet = QSqlSelectFree(QsqlConn)
                '--- 조회된 자료로 Insert처리
                SqlStr = " Insert into LAB04_DB..DJD010M ( " _
                        & " LABDATE, NUMGBN,  LABSQNO, ORDCD,  SUBCD, " _
                        & " RSTVAL,  REFVAL,  PANVAL,  DELVAL, REQGBN, " _
                        & " SPCGBN,  RETGBN,  RTNCD,   IDNO,   RSTID, " _
                        & " RSTDATE, ORDERNO, SYSDATE, SYSTIME ) values ( " _
                        & "'" & Left(sLabNo, 8) & "', " _
                        & "'" & Mid(sLabNo, 9, 1) & "', " _
                        & "'" & Right(sLabNo, 5) & "', " _
                        & "'" & .OrdCd & "', " _
                        & "'" & .SubNo & "', " _
                        & "'" & .Result & "', " _
                        & "'" & .Ref & "', " _
                        & "'', '',"
                For ii = 1 To 5
                    SqlStr = SqlStr & "'" & INSDATA(ii) & "', "
                Next ii
                SqlStr = SqlStr & "'" & D0COM_USERID & "', " _
                        & "'" & Format(Now, "YYYYMMDD") & "', " _
                        & "'" & sOrdNo & "', " _
                        & "'" & Format(Now, "YYYYMMDD") & "', " _
                        & "'" & Format(Now, "HHMMSS") & "') "
                        
                If QSqlDBExec(SqlStr, QsqlConn) <> QSQL_SUCCESS Then
                    Append_To_Server = False
                    Exit Function
                End If
                '---------------
            End If
        End If
        
        '--- 초기화
        .OrdCd = ""
        .SubNo = ""
        .Result = ""
        .Ref = ""
        '-----------
    End With
    
End Function

'
'   참고치 판정
'
Private Function Chk_Ref(sOrdCd As String, sSubNo As String, sRes As String, _
                        sex As String) As String

    Dim sStr    As String
    Dim SData() As String
    Dim iRet_Cd As Integer
    
    Dim LowVal  As Single
    Dim HighVal As Single
    Dim RefVal  As Single
    Dim RefChar As String

    Chk_Ref = ""

    If sex = "M" Then
        SqlStr = " Select REFLOM, REFHIM, REFCHAR, REFCHK "
    ElseIf sex = "F" Then
        SqlStr = " Select REFLOF, REFHIF, REFCHAR, REFCHK "
    End If
    SqlStr = SqlStr & "  From LAB01_DB..DJA060M " _
            & " where ORDCD = '" & sOrdCd & "'" _
            & "   and SUBCD = '" & sSubNo & "'"
    
    If QSqlDBExec(SqlStr, QsqlConn) = QSQL_SUCCESS Then
        If QSqlGetRow(sStr, QsqlConn) = QSQL_SUCCESS Then
            QSqlGetField 4, sStr, SData()
            
            If Trim(SData(4)) = "C" Then    '참고치 문자
                RefChar = Trim(SData(3))
                If RefChar <> Trim(sRes) Then
                    Chk_Ref = "*"
                End If
            ElseIf Trim(SData(4)) = "N" Then        '숫자
                RefVal = CSng(Trim(sRes))
                LowVal = CSng(SData(1)): HighVal = CSng(SData(2))
            
                If RefVal > HighVal Then
                    Chk_Ref = "H"
                ElseIf RefVal < LowVal Then
                    Chk_Ref = "L"
                End If
            End If
            
        End If
    End If
    iRet_Cd = QSqlSelectFree(QsqlConn)
    
End Function

'
'   해당 Sample No 의 결과 표시
'
Sub DIsp_Result(sSamNo As String, resRow As Integer)
           
    Dim ResData As Recordset
    Dim ResSql  As String
    Dim temp    As Variant
    Dim spdCol  As Integer
    
    Dim ii  As Integer
    
    ResSql = " SELECT * " _
            & "  FROM " & ResDB_Name & "" _
            & " WHERE SampNo = '" & sSamNo & "'" _
            & " ORDER BY EqCode"
            
    Set ResData = Db.OpenRecordset(ResSql, dbOpenDynaset)
    
    If ResData.EOF And ResData.BOF Then
        Exit Sub
    End If
        
    ResData.MoveFirst
    Do
        If Trim(ResData!EqCode) <> "" Then
            For ii = 1 To MAX_NUM
                If Trim(ResData!EqCode) = Trim(TestNameTable(ii).EqCd) Then
                    spdCol = ii + 4
                    Call spdList.SetText(spdCol, resRow, Trim(ResData!Result))
                    
                    Exit For
                End If
            Next ii
        End If
        ResData.MoveNext
    Loop Until (ResData.EOF)
    
    ResData.Close
    
End Sub
Private Sub Disp_SpdHeader()

    Dim ii  As Integer
        
    For ii = 1 To MAX_NUM
        If Trim(TestNameTable(ii).Name) <> "" Then
            Call spdList.SetText(ii + 4, 0, Trim(Left(TestNameTable(ii).Name, 7)))
        Else
            Call spdList.SetText(ii + 4, 0, " ")
        End If
    Next ii

End Sub





'
'   환자 신상자료 조회 Query문
'
Private Sub Query_Data()
 
    Dim tData() As String
    Dim sStr    As String
    Dim sLabNo  As String
    
    With spdList
        '----- 환자정보 조회/표시
        SqlDoc = " Select Distinct C2.PATNM, C2.SEX, D1.ORDERNO " _
                & "  from LAB04_DB..DJD010M D1, LAB03_DB..DJC020M C2 " _
                & " where D1.LABDATE = '" & Left$(TableSam!Lab_ID, 8) & "'" _
                & "   and D1.NUMGBN  = '" & Mid$(TableSam!Lab_ID, 9, 1) & "'" _
                & "   and D1.LABSQNO = '" & Right$(TableSam!Lab_ID, 5) & "'" _
                & "   and SUBSTRING(D1.ORDERNO,1,8) = C2.ORDDATE " _
                & "   and SUBSTRING(D1.ORDERNO,9,2) = C2.DEPTCD " _
                & "   and SUBSTRING(D1.ORDERNO,11,5) = C2.SEQNO " _
                & "   and D1.IDNO = C2.IDNO "
                
        If QSqlDBExec(SqlDoc, QsqlConn) = QSQL_SUCCESS Then
            If QSqlGetRow(sStr, QsqlConn) = QSQL_SUCCESS Then
                
                QSqlGetField 3, sStr, tData()
                
                .MaxRows = .MaxRows + 1
                
                sLabNo = Left$(TableSam!Lab_ID, 8) & "-" & Mid$(TableSam!Lab_ID, 9, 1) & "-" _
                        & Right$(TableSam!Lab_ID, 5)
                Call .SetText(2, .MaxRows, Trim(sLabNo))        '접수번호
                Call .SetText(3, .MaxRows, Trim(tData(1)))      '이름
                Call .SetText(4, .MaxRows, Trim$(TableSam!SampNo))  'Sample No
                Call .SetText(36, .MaxRows, Trim(tData(2)))     'Sex
                Call .SetText(37, .MaxRows, Trim(tData(3)))     'OrderNo
                '--- 결과 표시
                Call DIsp_Result(TableSam!SampNo, .MaxRows)
            End If
        End If
        Return_cd = QSqlSelectFree(QsqlConn)

    End With
    
End Sub


Private Sub Update_DB_Result(sSamNo As String, sEqCd As String, sResult As String, SeqN As Integer)

    Dim ResDB   As Database
    Dim ResTB   As Recordset
    
    Set ResDB = OpenDatabase(ResDb_Path & sEq_Name & "_" & Format(mskDate, "0000") & ".mdb")
    Set ResTB = ResDB.OpenRecordset(ResDB_Name)

    If RecordExists(ResTB, "PrimaryKey", sSamNo, sEqCd) Then
        ResTB.Edit
    Else
        ResTB.AddNew
        ResTB!SampNo = sSamNo
        ResTB!EqCode = sEqCd
    End If
    ResTB!SeqNo = Format(SeqN)
    ResTB!Result = sResult

    ResTB.Update
    ResTB.MoveLast
    
    ResTB.Close
    ResDB.Close
    
End Sub

Private Sub Update_DJC020M(OrdSqNo As String)

    Dim iRet_Cd As Integer
    Dim sStr    As String
    Dim tData() As String
    
    SqlStr = " Select count(*) from LAB03_DB..DJC050M " _
            & " where ORDDATE = '" & Left(OrdSqNo, 8) & "'" _
            & "   and DEPTCD = '" & Mid(OrdSqNo, 10, 2) & "'" _
            & "   and SEQNO = '" & Right(OrdSqNo, 5) & "'" _
            & "   and RSTGBN <> 'Y' "
    
    If QSqlDBExec(SqlStr, QsqlConn) = QSQL_SUCCESS Then
        If QSqlGetRow(sStr, QsqlConn) = QSQL_SUCCESS Then
            QSqlGetField 1, sStr, tData()
            
            If Val(tData(1)) = 0 Then
                SqlDoc = " Update LAB03_DB..DJC020M " _
                        & "   set ORDSTAT = '1' " _
                        & " where ORDDATE = '" & Left(OrdSqNo, 8) & "'" _
                        & "   and DEPTCD = '" & Mid(OrdSqNo, 10, 2) & "'" _
                        & "   and SEQNO = '" & Right(OrdSqNo, 5) & "'"
                iRet_Cd = QSqlDBExec(SqlDoc, QsqlCode)
            End If
        End If
    End If
    iRet_Cd = QSqlSelectFree(QsqlConn)
                
End Sub

Private Sub cboServer_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        txtLabID.SetFocus
    End If
    
End Sub

Private Sub cmdAppend_Click()
    
    Dim iR  As Integer
    Dim iC  As Integer
    Dim Tmp As Variant
    Dim ix1 As Integer
        
    Dim LabNo   As String
    Dim SampNo  As String
    Dim sSex    As String
    Dim sOrderNo    As String
    
    Dim ChkTrans    As Integer
    Dim ChkExist    As Integer
   
    If spdList.MaxRows = 0 Then
        MsgBox "등록할 자료가 존재하지 않습니다.", vbExclamation
        Exit Sub
    End If
    
    '--- Index Open
    If S0SUB_Open(D0COM_SERVER01, Me.hwnd, QsqlConn) <> QSQL_SUCCESS Then Exit Sub
    If S0SUB_Open(D0COM_SERVER01, Me.hwnd, QsqlCode) <> QSQL_SUCCESS Then Exit Sub
    
    MousePointer = 11
        
    ChkExist = False
    '----- Server로 결과등록
    For iR = 1 To spdList.MaxRows
        '--- Check Box
        Call spdList.GetText(1, iR, Tmp)
        If Tmp = True Then
            With spdList
                Call .GetText(2, iR, Tmp)
                LabNo = Trim(Tmp)           '접수번호
                Call .GetText(4, iR, Tmp)
                SampNo = Trim(Tmp)          'Sample No
                Call .GetText(36, iR, Tmp): sSex = Trim(Tmp)
                Call .GetText(37, iR, Tmp): sOrderNo = Trim(Tmp)
            End With
                        
            pnlStatus = "접수번호 " & LabNo & " 를 등록중입니다."
            ChkExist = True
            
            For iC = 1 To MAX_NUM
                If Trim(TestNameTable(iC).Code) <> "" Then
                    Call spdList.GetText(iC + 4, iR, Tmp)
                    With Insert_Server(iC)
                        .OrdCd = TestNameTable(iC).Code
                        .SubNo = ""
                        .Result = Trim(Tmp)
                        .Ref = Chk_Ref(.OrdCd, .SubNo, .Result, sSex)
                        '--- Hi_Result 내용 Update
                        Call Update_DB_Result(SampNo, TestNameTable(iC).EqCd, .Result, iC)
                    End With
                End If
            Next iC
                                
            '----- 구조체에 저장된 결과 Server에 등록
            ret = QSqlBeginTrans()
            DBEngine.Workspaces(0).BeginTrans
            ChkTrans = True
            
            For ix1 = 1 To MAX_NUM
                '----- 검사항목별 결과입력(Batch)
                If Append_To_Server(LabNo, ix1, sOrderNo) <> True Then
                    ChkTrans = False
                    Exit For
                End If
            Next ix1
            
            If ChkTrans = False Then
                DBEngine.Workspaces(0).Rollback
                ret = QSqlRollBack()      'TRANSACTION 에러종료
                pnlStatus = ""
            Else
                DBEngine.Workspaces(0).CommitTrans
                ret = QSqlCommitTrans()    'TRANSACTION 정상종료
                '--- 진료과별 처방내역 Update
'                Call Update_DJC020M(LabNo)
                Call Update_DJC020M(sOrderNo)
                
                '--- 등록체크 Update(MDB)
                Call Update_RegChk(SampNo)
                
                '--- 입력된 Row 색상반전
                With spdList
                    .Row = iR
                    .Col = -1
                    .BackColorStyle = BackColorStyleUnderGrid
                    .BackColor = &HC0FFC0
                End With
                '-------------------
                pnlStatus = "성공적으로 등록하였습니다."
            End If
        End If
    Next iR
    
    '--- Index Close
    ret = Qsqlclose(QsqlConn, ONECLOSE)
    ret = Qsqlclose(QsqlCode, ONECLOSE)

    If ChkExist <> True Then
        MsgBox "선택된 자료가 없습니다. 등록을 원하는 환자를 선택해 주십시오.", vbInformation
    End If
    
    MousePointer = 0
    
End Sub
Private Sub Update_RegChk(sSam As String)
                
    Dim SamDb   As Database
    Dim SamTb   As Recordset
    
    Set SamDb = OpenDatabase(ResDb_Path & sEq_Name & "_" & Format(mskDate, "0000") & ".mdb")
    Set SamTb = SamDb.OpenRecordset(SamDB_Name)

    SamTb.MoveFirst
    If RecordExist(SamTb, "PrimaryKey", sSam) Then
        SamTb.Edit
        SamTb!RegChk = "*"
        SamTb.Update
    End If

    SamTb.Close
    SamDb.Close
    
End Sub

Private Sub Add_Db_Result(SampNo As String, EqCode As String, Result As String)

    If RecordExists(TbResult, "PrimaryKey", SampNo, EqCode) Then
        TbResult.Edit
    Else
        TbResult.AddNew
        TbResult!SampNo = SampNo
        TbResult!EqCode = EqCode
    End If
    TbResult!Result = Result

    TbResult.Update
    TbResult.MoveLast

End Sub



Private Sub cmdExit_Click()

    Unload Me
    
End Sub




Private Sub CmdQuery_Click()
        
    Dim FileName    As String
    Dim iRet    As Integer
    
    pnlStatus = ""
    
    If Trim(mskDate) = "" Then Exit Sub
    
    FileName = mskDate
    
    '----- 결과화일 존재여부 검사
    If iFileExists(ResDb_Path & sEq_Name & "_" & FileName & ".mdb") <> True Then
        MsgBox "화일 " & sEq_Name & "_" & FileName & ".mdb 가 존재하지 않습니다.", _
                vbCritical, Me.Caption
        mskDate.SetFocus
        Exit Sub
    End If
    
    Set Db = OpenDatabase(ResDb_Path + sEq_Name & "_" + FileName + ".mdb")

    Set TableVar = Db.OpenRecordset(ResDB_Name)
    Set TableSam = Db.OpenRecordset(SamDB_Name)

    If TableSam.BOF And TableSam.EOF Then
        MsgBox "수신된 결과가 존재하지 않습니다. 결과받기를 먼저 실행해 주십시오.", _
                vbExclamation, Me.Caption
        GoTo DB_Close
    Else
        spdList.MaxRows = 0
    End If
        
    TableSam.Index = "PrimaryKey"
    TableSam.MoveFirst
    
    '--- Index Open
    If S0SUB_Open(D0COM_SERVER01, Me.hwnd, QsqlConn) <> QSQL_SUCCESS Then GoTo DB_Close
    
    Do Until TableSam.EOF
        '----- Server에 등록 또는 미등록 자료만 조회
        If cboServer.ListIndex = 0 Then
            If TableSam!RegChk <> "*" Or IsNull(TableSam!RegChk) Then       '미등록자료만 조회
                Call Query_Data     'Query/Display 문
            End If
        ElseIf cboServer.ListIndex = 1 Then
            If TableSam!RegChk = "*" Then
                Call Query_Data
            End If
        End If
        TableSam.MoveNext
        '-------------------------------------------
    Loop
    
'    pnlCnt = TableSam.RecordCount   'Total Count
    
    '--- Index Close
    iRet = Qsqlclose(QsqlConn, ONECLOSE)
    
DB_Close:
    TableSam.Close
    TableVar.Close
    Db.Close
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyEscape
            Call cmdExit.DoClick
    End Select

End Sub

Private Sub Form_Load()
        
    'form을 위치
    Me.Top = 0
    Me.Left = 0
    Me.Height = Hi_Main.Height - Hi_Main.pnlBack.Height - 500
    Me.Width = Hi_Main.Width - 200
    '-----------
    MousePointer = 11
        
    ResDB_Name = sEq_Name & "_Result"
    SamDB_Name = sEq_Name & "_Sample"

    mskDate = Format(Now, "mmdd")
    
    '--- Combo Box Display
    With cboServer
        .AddItem ("미등록된 결과자료")
        .AddItem ("등록된 결과자료")
        '조회구분별 환자 신상자료 조회/화면 표시
        .ListIndex = 0
    End With
    
    Call Disp_SpdHeader     'Spread Header 검사항목 표시
    
    MousePointer = 0
    
End Sub

Private Sub mskDate_GotFocus()

    With mskDate
        .SelStart = 0
        .SelLength = .MaxLength
    End With
    
End Sub


Private Sub mskDate_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        cboServer.SetFocus
    End If
    
End Sub


Private Sub spdList_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    Dim sNo As Long
    Dim eNo As Long
    Dim irow    As Integer
    Dim Tmp As Variant
    
    If BlockRow = 0 Or BlockRow2 = 0 Or BlockRow = BlockRow2 Then Exit Sub
    
    If BlockRow < BlockRow2 Then
        sNo = BlockRow: eNo = BlockRow2
    Else
        sNo = BlockRow2: eNo = BlockRow
    End If

    For irow = sNo To eNo
        With spdList
            Call .GetText(1, irow, Tmp)
            If Tmp = True Then
                Call .SetText(1, irow, "0")
            Else
                Call .SetText(1, irow, "1")
            End If
        End With
    Next irow
    
End Sub


Private Sub spdList_GotFocus()

    With spdList
        If OldRow <> 0 Then
            .Row = OldRow
            .Col = -1
            .BackColor = &H80000005
        End If
        If .ActiveRow = 1 Then
            .Row = .ActiveRow
            
            .BackColor = &HEEEFFF
            OldRow = .ActiveRow
        End If
    End With
        
End Sub

Private Sub spdList_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
    If Row <> NewRow Then
        If OldRow = 0 Then OldRow = Row
        
        With spdList
            If NewRow <> -1 Then
                .Row = OldRow
                .Col = -1
                .BackColor = &H80000005
                
                .Row = NewRow
                .Col = -1

                .BackColor = &HEEEFFF
                OldRow = NewRow
            End If
        End With
    End If
    
End Sub



Private Sub txtLabID_GotFocus()

    With txtLabID
        .SelStart = 0
        .SelLength = .MaxLength
        .Tag = .Text
    End With
    
End Sub


Private Sub txtLabID_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        spdList.SetFocus
    End If
    
End Sub


Private Sub txtLabID_LostFocus()
    
    Dim ii  As Integer
    Dim Tmp As Variant

    If Trim(txtLabID.Text) <> Trim(txtLabID.Tag) Then
        If Trim(mskDate) = "" Then Exit Sub
        
        Set Db = OpenDatabase(ResDb_Path & sEq_Name & "_" & Format(mskDate, "0000") & ".mdb")
        Set TableSam = Db.OpenRecordset(SamDB_Name)
        
        If TableSam.EOF And TableSam.BOF Then
            TableSam.Close
            Db.Close
        End If
        
        With spdList
            If .MaxRows <> 0 Then
                For ii = 1 To .MaxRows
                    Call .GetText(2, ii, Tmp)
                    If Left(Tmp, 8) & Mid(Tmp, 10, 1) & Right(Tmp, 5) = Trim(TableSam!Lab_ID) Then
                        .TopRow = ii
                        Exit For
                    End If
                Next ii
            Else
                Call Query_Data
            End If
        End With
        
        TableSam.Close
        Db.Close
    End If
    
End Sub

