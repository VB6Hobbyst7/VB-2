VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmWorkList 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "워크리스트 조회"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17205
   Icon            =   "frmWorkList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   17205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin FPSpread.vaSpread vasPrint 
      Height          =   3315
      Left            =   3120
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   11580
      _Version        =   393216
      _ExtentX        =   20426
      _ExtentY        =   5847
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   9
      SpreadDesigner  =   "frmWorkList.frx":000C
   End
   Begin VB.CommandButton cmdWorkPrint 
      Caption         =   "출력"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   12
      Top             =   180
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "Check1"
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   900
      TabIndex        =   10
      Top             =   810
      Width           =   225
   End
   Begin VB.TextBox txtSeq 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6690
      TabIndex        =   8
      Text            =   "1"
      Top             =   210
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   7
      Top             =   180
      Width           =   1395
   End
   Begin VB.CommandButton cmdDownClose 
      Caption         =   "Down >> Close"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   6
      Top             =   180
      Width           =   1395
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "조회"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4710
      TabIndex        =   1
      Top             =   180
      Width           =   1395
   End
   Begin VB.CommandButton cmdDownLoad 
      Caption         =   "Down"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   0
      Top             =   180
      Width           =   1395
   End
   Begin MSComCtl2.DTPicker dtpStartDt 
      Height          =   315
      Left            =   1290
      TabIndex        =   2
      Top             =   180
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   63504385
      CurrentDate     =   40457
   End
   Begin MSComCtl2.DTPicker dtpStopDt 
      Height          =   315
      Left            =   3030
      TabIndex        =   3
      Top             =   180
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   63504385
      CurrentDate     =   40457
   End
   Begin FPSpread.vaSpread vasWorkList 
      Height          =   5865
      Left            =   300
      TabIndex        =   11
      Top             =   720
      Width           =   16695
      _Version        =   393216
      _ExtentX        =   29448
      _ExtentY        =   10345
      _StockProps     =   64
      ButtonDrawMode  =   4
      ColHeaderDisplay=   0
      ColsFrozen      =   16
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   17
      MaxRows         =   20
      MoveActiveOnFocus=   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmWorkList.frx":1A98
   End
   Begin VB.Label Label2 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "Seq"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6180
      TabIndex        =   9
      Top             =   270
      Width           =   375
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "조회일자"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   330
      TabIndex        =   5
      Top             =   270
      Width           =   780
   End
   Begin VB.Label Label7 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2820
      TabIndex        =   4
      Top             =   270
      Width           =   105
   End
End
Attribute VB_Name = "frmWorkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAll_Click()
    Dim iRow As Long
    
    If chkAll.Value = 1 Then
        For iRow = 1 To vasWorkList.DataRowCnt
            vasWorkList.Row = iRow
            vasWorkList.Col = 1
            
            vasWorkList.Value = 1
        Next iRow
    ElseIf chkAll.Value = 0 Then
        For iRow = 1 To vasWorkList.DataRowCnt
            vasWorkList.Row = iRow
            vasWorkList.Col = 1
            
            vasWorkList.Value = 0
        Next iRow
    End If
    
End Sub

Private Sub cmdClose_Click()
    
    Unload Me
    
End Sub

Private Sub cmdDownClose_Click()
    
    Call cmdDownLoad_Click
    
    Call cmdClose_Click
    
End Sub

Private Sub cmdDownLoad_Click()
    Dim intVasRow As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    Dim j  As Integer
    
    j = 0
    With vasWorkList
        For intRow = 1 To .MaxRows
            .Row = intRow
            .Col = colCheckBox
            If .Value = 1 Then
                frmInterface.vasID.MaxRows = frmInterface.vasID.MaxRows + 1
                DoEvents
                intVasRow = frmInterface.vasID.MaxRows
                
                If GetText(vasWorkList, intRow, colBARCODE) = "" Then
                    frmInterface.vasID.MaxRows = frmInterface.vasID.MaxRows - 1
                    Exit Sub
                End If
                
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colSpecNo), intVasRow, colSpecNo)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colCheckBox), intVasRow, colCheckBox)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colHOSPDATE), intVasRow, colHOSPDATE)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colBARCODE), intVasRow, colBARCODE)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colCHARTNO), intVasRow, colCHARTNO)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colPID), intVasRow, colPID)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colINOUT), intVasRow, colINOUT)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colPNAME), intVasRow, colPNAME)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colPSEX), intVasRow, colPSEX)
                Call SetText(frmInterface.vasID, GetText(vasWorkList, intRow, colPAGE), intVasRow, colPAGE)
                
               ' frmInterface.txtNum = frmInterface.txtNum + 1
                
                For intCol = colState + 1 To .MaxCols
                    .Row = intVasRow
                    .Col = intCol
                    If .BackColor = vbYellow Then
                        Call SetBackColor(frmInterface.vasID, intVasRow, intVasRow, intCol, intCol, 255, 255, 0)
                    End If
                Next
        
                
                .Col = 1
                .Value = "0"
            End If
        Next
        frmInterface.vasID.RowHeight(-1) = 12
    End With

    frmInterface.dtpStartDt.Value = dtpStartDt.Value
    frmInterface.dtpStopDt.Value = dtpStopDt.Value
'    frmInterface.txtSeq.Text = txtSeq.Text



'    Dim i As Integer
'
'    If KeyAscii = vbKeyReturn Then
'        For i = 1 To vasWorkList.MaxRows
'            vasWorkList.Row = i
'            vasWorkList.Col = 1
'            If vasWorkList.Value = "1" Then
'                If Trim(txtPos.Text) = "" Then
'                    txtPos.Text = "1"
'                End If
'                Call SetText(frmInterface.vasworklist, Format(txtPos.Text, "0000"), i, 0)
'                txtPos.Text = Format(txtPos.Text + 1, "0000")
'            End If
'        Next
'    End If
End Sub

Private Sub cmdSearch_Click()

    Call GetWorkList_AMIS(Format(dtpStartDt, "yyyymmdd"), Format(dtpStopDt, "yyyymmdd"))
    
End Sub

Private Sub GetWorkList_MCC(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim strBarcode    As String
    Dim blnSame     As Boolean
    
    If pBarNo = "" Then
        vasWorkList.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasWorkList.ReDraw = False
    

'''          SQL = "SELECT DISTINCT ORD_YMD, BCODE_NO, RECEPT_NO, PTNT_NO,PTNT_NM,AGE,SEX,ORD_CD" & vbCr
'''    SQL = SQL & "  FROM MCCSI.H7LIS_BCODE_ORD " & vbCr
'''    SQL = SQL & " WHERE ORD_YMD between '" & pFrDt & "' AND '" & pToDt & "'" & vbCr
'''    SQL = SQL & "   AND ORD_CD IN (" & gAllExam & ") " & vbCr
'''    SQL = SQL & "   AND RESULT_TYPE = '20'" & vbLf & vbCr
'''    SQL = SQL & "  ORDER BY ORD_YMD,RECEPT_NO,BCODE_NO "

          SQL = "SELECT DISTINCT  a.ptnt_no, c.ptnt_nm, a.recept_no, a.spc_cd, " & vbCr
    SQL = SQL & " (select codeval1 from pm_mst_div_key1 where codediv = 'LAB01' and codekey1 = a.spc_cd) as spc_nm " & vbCr
    SQL = SQL & "      , a.sts_cd, a.acc_ymd, a.ord_cd " & vbCr
    SQL = SQL & "  FROM h3lab_result a, h1opdin b, hz_mst_ptnt c " & vbCr
    SQL = SQL & " WHERE a.recept_no = b.recept_no " & vbCr
    SQL = SQL & "   AND a.sutak_cd = ''" & vbCr
    SQL = SQL & "   AND a.ptnt_no  = c.ptnt_no" & vbCr
    SQL = SQL & "   AND a.sts_cd   = 'A'" & vbCr                                                               ' A:접수 R:결과"
    SQL = SQL & "   AND a.acc_ymd between '" & pFrDt & "' AND '" & pToDt & "'" & vbCr
    SQL = SQL & "   AND a.ord_cd IN (" & gAllExam & ") " & vbCr
    SQL = SQL & " Order by recept_no " & vbCr

    Call SetSQLData("워크조회", SQL)

    '-- Record Count 가져옴
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasWorkList
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasWorkList, i, colHOSPDATE)
                    strBarcode = GetText(vasWorkList, i, colBARCODE)
                    If Trim(RS("acc_ymd")) = strDate And Trim(RS("recept_no")) = strBarcode Then
                        blnSame = True
                    End If
                    
                    For intCol = colState + 1 To vasWorkList.MaxCols
                        If Trim(RS.Fields("ord_cd")) = gArrEquip(intCol - colState, 3) Then
                            vasWorkList.Row = .MaxRows
                            vasWorkList.Col = intCol
                            vasWorkList.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    
                    SetText vasWorkList, txtSeq.Text, .MaxRows, colSpecNo
                    SetText vasWorkList, "1", .MaxRows, colCheckBox
                    SetText vasWorkList, Trim(RS.Fields("acc_ymd")) & "", .MaxRows, colHOSPDATE
                    SetText vasWorkList, Trim(RS.Fields("recept_no")) & "", .MaxRows, colBARCODE
                    'SetText vasWorkList, Trim(RS.Fields("ptnt_no")) & "", .MaxRows, colCHARTNO
                    SetText vasWorkList, Trim(RS.Fields("ptnt_no")) & "", .MaxRows, colPID
                    SetText vasWorkList, Trim(RS.Fields("ptnt_nm")) & "", .MaxRows, colPNAME
                    'SetText vasWorkList, Trim(RS.Fields("AGE")) & "", .MaxRows, colPAGE
                    'SetText vasWorkList, Trim(RS.Fields("SEX")) & "", .MaxRows, colPSEX
                    
                    txtSeq.Text = txtSeq.Text + 1
                    
                    For intCol = colState + 1 To vasWorkList.MaxCols
                        If Trim(RS.Fields("ord_cd")) = gArrEquip(intCol - colState, 3) Then
                            vasWorkList.Row = .MaxRows
                            vasWorkList.Col = intCol
                            vasWorkList.BackColor = vbYellow
                            Exit For
                        End If
                    Next
                End If
                blnSame = False
            End With
                        
            RS.MoveNext
        Loop
        chkAll.Value = "1"
    Else
        'StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
        chkAll.Value = "0"
    End If
    
    RS.Close
    
    '-- 프로그레스바 닫기
    Unload frmProgress
    
    vasWorkList.RowHeight(-1) = 12
    vasWorkList.ReDraw = True
    
End Sub

Private Sub GetWorkList_EHEALTH(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim strBarcode    As String
    Dim blnSame     As Boolean
    
    If pBarNo = "" Then
        vasWorkList.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasWorkList.ReDraw = False
    
          SQL = "SELECT DISTINCT b.OBODORDT"    '입력일
    SQL = SQL & ", a.APATMRNO"                  '등록번호
    SQL = SQL & ", b.OBODCASE"                  '내원번호
    SQL = SQL & ", b.OBODORNO"                  'ORDER NUMBER
    SQL = SQL & ", b.OBODORSQ"                  'ORDER SEQUENCE
    SQL = SQL & ", b.OBODIOGB"                  '입/외 I=입원/O=외래
    SQL = SQL & ", a.APATNAME"                  '환자성명
    SQL = SQL & ", a.APATPSEX"                  '성별(M/F)
    SQL = SQL & ", a.APATJMN1"                  '주민번호(년월일)
    SQL = SQL & ", b.OBODCODE"                  '수가코드
    SQL = SQL & ", c.OBSUCODE"                  '검사코드
    SQL = SQL & ", c.OBSUSUBC" & vbCr           '검사코드SUB
    
    SQL = SQL & "  FROM ABPATMST a"             '환자기본정보
    SQL = SQL & "      ,OBODRMTM b"            '처방내역 Table
    SQL = SQL & "      ,OBSURSTM c " & vbCr     '검사결과(수치결과) Table
    
    SQL = SQL & " WHERE a.APATMRNO = b.OBODMRNO " & vbCr                                '등록번호,고객번호
    SQL = SQL & "   AND a.APATMRNO = c.OBSUMRNO " & vbCr                                '등록번호,고객번호
    SQL = SQL & "   AND b.OBODCASE = c.OBSUCASE " & vbCr                                '내원번호
    SQL = SQL & "   AND b.OBODORNO = c.OBSUORNO " & vbCr                                'ORDER NUMBER
    SQL = SQL & "   AND b.OBODORSQ = c.OBSUORSQ " & vbCr                                'ORDER SEQUENCE
    SQL = SQL & "   AND (c.OBSURSLT IS NULL OR c.OBSURSLT = '')" & vbCr                 '검사결과
    SQL = SQL & "   AND b.OBODORDT between '" & pFrDt & "' AND '" & pToDt & "'" & vbCr  '입력일
    SQL = SQL & "   AND RTRIM(c.OBSUCODE) + '|' + RTRIM(c.OBSUSUBC) IN (" & gAllExam & ") " & vbCr    '검사코드 + '|' + OBSUSUBC
    SQL = SQL & "   AND b.OBODSTAT = 'AC' " & vbCr                                      '필수 기본 = 'OE', 채혈시 = 'AC'
    SQL = SQL & " Order by b.OBODORDT,a.APATMRNO,b.OBODORNO,b.OBODORSQ " & vbCr

    Call SetSQLData("워크조회", SQL)

    '-- Record Count 가져옴
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasWorkList
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasWorkList, i, colHOSPDATE)
                    strBarcode = GetText(vasWorkList, i, colBARCODE)
                    If Trim(RS("OBODORDT")) = strDate And Trim(RS("APATMRNO")) = strBarcode Then
                        blnSame = True
                    End If
                    
                    For intCol = colState + 1 To vasWorkList.MaxCols
                        If Trim(RS.Fields("OBSUCODE")) & "|" & Trim(RS.Fields("OBSUSUBC")) = gArrEquip(intCol - colState, 3) Then
                            vasWorkList.Row = .MaxRows
                            vasWorkList.Col = intCol
                            vasWorkList.BackColor = vbYellow
                            vasWorkList.Text = "◆"
                            Exit For
                        End If
                    Next
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    
                    SetText vasWorkList, txtSeq.Text, .MaxRows, colSpecNo
                    SetText vasWorkList, "1", .MaxRows, colCheckBox
                    SetText vasWorkList, Trim(RS.Fields("OBODORDT")) & "", .MaxRows, colHOSPDATE
                    SetText vasWorkList, Trim(RS.Fields("APATMRNO")) & "", .MaxRows, colBARCODE
                    SetText vasWorkList, Trim(RS.Fields("OBODCASE")) & "", .MaxRows, colCHARTNO
                    SetText vasWorkList, Trim(RS.Fields("OBODORNO")) & "", .MaxRows, colPID
                    SetText vasWorkList, IIf(Trim(RS.Fields("OBODIOGB")) & "" = "I", "입원", "외래"), .MaxRows, colINOUT
                    SetText vasWorkList, Trim(RS.Fields("APATNAME")) & "", .MaxRows, colPNAME
                    SetText vasWorkList, Trim(RS.Fields("APATPSEX")) & "", .MaxRows, colPSEX
                    SetText vasWorkList, Trim(RS.Fields("APATJMN1")) & "", .MaxRows, colPAGE
                    
                    txtSeq.Text = txtSeq.Text + 1
                    
                    For intCol = colState + 1 To vasWorkList.MaxCols
                        If Trim(RS.Fields("OBSUCODE")) & "|" & Trim(RS.Fields("OBSUSUBC")) = gArrEquip(intCol - colState, 3) Then
                            vasWorkList.Row = .MaxRows
                            vasWorkList.Col = intCol
                            vasWorkList.BackColor = vbYellow
                            vasWorkList.Text = "◆"
                            Exit For
                        End If
                    Next
                End If
                blnSame = False
            End With
                        
            RS.MoveNext
        Loop
        chkAll.Value = "1"
    Else
        'StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
        chkAll.Value = "0"
    End If
    
    RS.Close
    
    '-- 프로그레스바 닫기
    'Unload frmProgress
    
    vasWorkList.RowHeight(-1) = 12
    vasWorkList.ReDraw = True
    
End Sub

Private Sub GetWorkList_AMIS(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim strBarcode    As String
    Dim blnSame     As Boolean
    
    If pBarNo = "" Then
        vasWorkList.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    vasWorkList.ReDraw = False
    
'          SQL = "SELECT DISTINCT b.OBODORDT"    '입력일
'    SQL = SQL & ", a.APATMRNO"                  '등록번호
'    SQL = SQL & ", b.OBODCASE"                  '내원번호
'    SQL = SQL & ", b.OBODORNO"                  'ORDER NUMBER
'    SQL = SQL & ", b.OBODORSQ"                  'ORDER SEQUENCE
'    SQL = SQL & ", b.OBODIOGB"                  '입/외 I=입원/O=외래
'    SQL = SQL & ", a.APATNAME"                  '환자성명
'    SQL = SQL & ", a.APATPSEX"                  '성별(M/F)
'    SQL = SQL & ", a.APATJMN1"                  '주민번호(년월일)
'    SQL = SQL & ", b.OBODCODE"                  '수가코드
'    SQL = SQL & ", c.OBSUCODE"                  '검사코드
'    SQL = SQL & ", c.OBSUSUBC" & vbCr           '검사코드SUB
'
'    SQL = SQL & "  FROM ABPATMST a"             '환자기본정보
'    SQL = SQL & "      ,OBODRMTM b"            '처방내역 Table
'    SQL = SQL & "      ,OBSURSTM c " & vbCr     '검사결과(수치결과) Table
'
'    SQL = SQL & " WHERE a.APATMRNO = b.OBODMRNO " & vbCr                                '등록번호,고객번호
'    SQL = SQL & "   AND a.APATMRNO = c.OBSUMRNO " & vbCr                                '등록번호,고객번호
'    SQL = SQL & "   AND b.OBODCASE = c.OBSUCASE " & vbCr                                '내원번호
'    SQL = SQL & "   AND b.OBODORNO = c.OBSUORNO " & vbCr                                'ORDER NUMBER
'    SQL = SQL & "   AND b.OBODORSQ = c.OBSUORSQ " & vbCr                                'ORDER SEQUENCE
'    SQL = SQL & "   AND (c.OBSURSLT IS NULL OR c.OBSURSLT = '')" & vbCr                 '검사결과
'    SQL = SQL & "   AND b.OBODORDT between '" & pFrDt & "' AND '" & pToDt & "'" & vbCr  '입력일
'    SQL = SQL & "   AND RTRIM(c.OBSUCODE) + '|' + RTRIM(c.OBSUSUBC) IN (" & gAllExam & ") " & vbCr    '검사코드 + '|' + OBSUSUBC
'    SQL = SQL & "   AND b.OBODSTAT = 'AC' " & vbCr                                      '필수 기본 = 'OE', 채혈시 = 'AC'
'    SQL = SQL & " Order by b.OBODORDT,a.APATMRNO,b.OBODORNO,b.OBODORSQ " & vbCr

    SQL = ""
    SQL = SQL & "SELECT P.PATID as 챠트번호, P.PATNAME as 이름, P.SEX, O.ACPTDATE as 접수일자, O.ACPTSEQ, O.RSVACPTSTATE, O.RESULTSTATE, O.DEPTCODE, O.ORDERDATE, O.SLIPNO, O.IOFLAG, O.ORDERCODE, O.ORDERNAME, R.SPCMNO as 바코드번호, R.RESULTFLAG, R.RESULTNO, R.RESULTITEMCODE as ITEM " & vbCr
    SQL = SQL & "  FROM sghis.registinfos O, sghis.resultofnum R, sghis.PATMST P " & vbCr
    SQL = SQL & " WHERE O.acptdate = R.acptdate " & vbCr
    SQL = SQL & "   AND O.acptdate between '" & pFrDt & "' and '" & pToDt & "'"
    SQL = SQL & "   AND R.resultitemcode in (" & gAllExam & ")"
    SQL = SQL & "   AND O.patid = R.patid " & vbCr
    SQL = SQL & "   AND O.acptseq = R.acptseq " & vbCr
    SQL = SQL & "   AND O.patid = P.patid " & vbCr
    SQL = SQL & "   AND O.CLAS = 4 " & vbCr '임상병리
    SQL = SQL & "   AND R.RESULTFLAG = 0 " & vbCr
    SQL = SQL & "   AND O.rsvacptstate < '5' " & vbCr
'    If chkGum.Value = "1" Then
'        SQL = SQL & "   AND O.ORDERCODE IN ('LU101','LU00011','LU00015') "
'    Else
'        SQL = SQL & "   AND O.ORDERCODE = 'LU101' "
'    End If
    SQL = SQL & "  ORDER BY R.SPCMNO"
    
    Call SetSQLData("워크조회", SQL)

    '-- Record Count 가져옴
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        
        Do Until RS.EOF
            iCnt = iCnt + 1
            With vasWorkList
                .ReDraw = False
                For i = 1 To .DataRowCnt
                    strDate = GetText(vasWorkList, i, colHOSPDATE)
                    strBarcode = GetText(vasWorkList, i, colBARCODE)
                    If Trim(RS("접수일자")) = strDate And Trim(RS("바코드번호")) = strBarcode Then
                        blnSame = True
                    End If
                    
                    For intCol = colState + 1 To vasWorkList.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasWorkList.Row = .MaxRows
                            vasWorkList.Col = intCol
                            vasWorkList.BackColor = vbYellow
                            vasWorkList.Text = "◆"
                            Exit For
                        End If
                    Next
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    
                    SetText vasWorkList, txtSeq.Text, .MaxRows, colSpecNo
                    SetText vasWorkList, "1", .MaxRows, colCheckBox
                    SetText vasWorkList, Trim(RS.Fields("접수일자")) & "", .MaxRows, colHOSPDATE
                    SetText vasWorkList, Trim(RS.Fields("바코드번호")) & "", .MaxRows, colBARCODE
                    SetText vasWorkList, Trim(RS.Fields("챠트번호")) & "", .MaxRows, colCHARTNO
                    SetText vasWorkList, Trim(RS.Fields("SLIPNO")) & "", .MaxRows, colPID
                    SetText vasWorkList, IIf(Trim(RS.Fields("IOFLAG")) & "" = "I", "입원", "외래"), .MaxRows, colINOUT
                    SetText vasWorkList, Trim(RS.Fields("이름")) & "", .MaxRows, colPNAME
                    SetText vasWorkList, Trim(RS.Fields("SEX")) & "", .MaxRows, colPSEX
                    
'                    SetText vasWorkList, "20170313", .MaxRows, colHOSPDATE
'                    SetText vasWorkList, "1234567890", .MaxRows, colBARCODE
'                    SetText vasWorkList, "987654", .MaxRows, colCHARTNO
'                    SetText vasWorkList, "1234", .MaxRows, colPID
'                    SetText vasWorkList, "입원", .MaxRows, colINOUT
'                    SetText vasWorkList, "홍길동", .MaxRows, colPNAME
'                    SetText vasWorkList, "M", .MaxRows, colPSEX
                    
                    txtSeq.Text = txtSeq.Text + 1
                    
                    For intCol = colState + 1 To vasWorkList.MaxCols
                        If Trim(RS.Fields("ITEM")) = gArrEquip(intCol - colState, 3) Then
                            vasWorkList.Row = .MaxRows
                            vasWorkList.Col = intCol
                            vasWorkList.BackColor = vbYellow
                            vasWorkList.Text = "◆"
                            Exit For
                        End If
                    Next
                End If
                blnSame = False
            End With
                        
            RS.MoveNext
        Loop
        chkAll.Value = "1"
    Else
        'StatusBar1.Panels(3).Text = "조회 대상자가 없습니다."
        chkAll.Value = "0"
    End If
    
    RS.Close
    
    vasWorkList.RowHeight(-1) = 12
    vasWorkList.ReDraw = True
    
End Sub


Private Sub cmdWorkPrint_Click()
Dim iRow As Integer
Dim iCol As Integer
Dim j As Integer

Dim sCurDate As String
Dim sSerDate As String
Dim sHead As String
Dim sFoot As String
Dim strTestNm As String

    ClearSpread vasPrint

    j = 1

    vasPrint.RowHeight(-1) = 20
    
    For iRow = 1 To vasWorkList.DataRowCnt
        vasWorkList.Row = iRow
        vasWorkList.Col = colCheckBox

        If vasWorkList.Value = 1 Then
            SetText vasPrint, Trim(GetText(vasWorkList, iRow, colBARCODE)), j, 1     '검체번호
            SetText vasPrint, Trim(GetText(vasWorkList, iRow, colCHARTNO)), j, 2     '환자번호
            SetText vasPrint, Trim(GetText(vasWorkList, iRow, colPNAME)), j, 3     '환자이름

            SetText vasPrint, Trim(GetText(vasWorkList, iRow, colPSEX)), j, 4     '성별
            SetText vasPrint, Trim(GetText(vasWorkList, iRow, colPAGE)), j, 5     '나이
            
            strTestNm = ""
            For iCol = colState + 1 To vasWorkList.MaxCols
                vasWorkList.Row = iRow
                vasWorkList.Col = iCol
                If vasWorkList.Text = "◆" Then
                    strTestNm = strTestNm & Trim(GetText(vasWorkList, 0, iCol)) & "/"
                End If
            Next
            SetText vasPrint, strTestNm, j, 6     '검사항목

            j = j + 1
        End If
    Next iRow

    If vasPrint.DataRowCnt < 1 Then
        MsgBox "출력할 자료가 없습니다.", , "알 림"
        Exit Sub
    End If

    sCurDate = GetDateFull

    sSerDate = Trim(dtpStartDt.Value) & " - " & Trim(dtpStopDt.Value)

    vasPrint.PrintOrientation = 1   ' SS_PRINTORIENT_PORTRAIT
    vasPrint.PrintAbortMsg = "인쇄중 입니다 ..."
    vasPrint.PrintJobName = "WorkList 출력"


    sHead = "/fn""궁서체"" /fz""12"" /fb1 /fi0 /fu0 " & "/c" & "▣ 혈액학 WorkList ▣" & "/n/n " & _
            "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 " & "/c" & "처방일자 : " & dtpStartDt & " ~ " & dtpStopDt

    sFoot = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""궁서체"" /fz""11"" /fb1 /fi0 /fu0 /r" & " 나무정원병원 검사실"

    vasPrint.PrintHeader = sHead
    vasPrint.PrintFooter = sFoot

    vasPrint.PrintMarginTop = 680
    vasPrint.PrintMarginBottom = 680
'현재 SS가 비대칭으로 출력함
'    vaslist.PrintMarginLeft = 720
    vasPrint.PrintMarginLeft = 0
    vasPrint.PrintMarginRight = 0

    vasPrint.PrintColor = True
    vasPrint.PrintGrid = True

'Set printing range
    vasPrint.PrintType = 0  'SS_PRINT_ALL(default)

    vasPrint.PrintShadows = True

    vasPrint.Action = 13 'SS_ACTION_PRINT

'    vasWorkList.PrintOrientation = PrintOrientationLandscape '가로출력
'    vasWorkList.Action = ActionPrint
    
End Sub

Private Sub Form_Load()
    
    dtpStartDt.Value = frmInterface.dtpStartDt.Value
    dtpStopDt.Value = frmInterface.dtpStopDt.Value
'    txtSeq.Text = frmInterface.txtSeq.Text
    
    vasWorkList.MaxRows = 0
    
End Sub



Private Sub txtSeq_KeyPress(KeyAscii As Integer)
    Dim intRow As Integer
    
    If KeyAscii = vbKeyReturn Then
        
        For intRow = vasWorkList.ActiveRow To vasWorkList.MaxRows
            Call SetText(vasWorkList, Val(txtSeq.Text), intRow, colSpecNo)
            txtSeq.Text = txtSeq.Text + 1
        Next
        
       ' txtSeq.Text = Format(txtSeq.Text, "0000")
    
    End If


End Sub

Private Sub vasWorkList_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim pGrid_Point As Integer
    Dim sBarcode As String
    Dim sChartNo As String
    
    If Row = 0 Then Exit Sub
    
    With vasWorkList
        '.Col = Col
        '.Row = Row
        '.Col = colBarcode
        pGrid_Point = SeqSearch(frmInterface.vasID, GetText(vasWorkList, Row, colBARCODE), colBARCODE)

        If pGrid_Point = 0 Then
            pGrid_Point = SeqNullSearch(frmInterface.vasID, Trim(.Text), colBARCODE)
            If pGrid_Point = 0 Then
                frmInterface.vasID.MaxRows = frmInterface.vasID.MaxRows + 1
                pGrid_Point = frmInterface.vasID.MaxRows
            End If
            .RowHeight(-1) = 12
        End If
        
'        .Row = Row: .Col = colBarcode
'        sBarcode = Trim(.Text)
        
        
'        Call frmInterface.vasworklist.SetText(colSpecNo, pGrid_Point, "1")
'        Call frmInterface.vasworklist.SetText(1, pGrid_Point, "1")
'        Call frmInterface.vasworklist.SetText(4, pGrid_Point, .Text)

        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colSpecNo), pGrid_Point, colSpecNo)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colCheckBox), pGrid_Point, colCheckBox)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colHOSPDATE), pGrid_Point, colHOSPDATE)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colBARCODE), pGrid_Point, colBARCODE)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colPID), pGrid_Point, colPID)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colCHARTNO), pGrid_Point, colCHARTNO)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colPNAME), pGrid_Point, colPNAME)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colPSEX), pGrid_Point, colPSEX)
        Call SetText(frmInterface.vasID, GetText(vasWorkList, Row, colPAGE), pGrid_Point, colPAGE)



'        .Row = Row: .Col = 5
'        Call vasworklist.SetText(5, pGrid_Point, .Text)
'        .Row = Row: .Col = 6
'        Call vasworklist.SetText(6, pGrid_Point, .Text)
'        .Row = Row: .Col = 7
'        Call vasworklist.SetText(7, pGrid_Point, .Text)
'        .Row = Row: .Col = 8
'        Call vasworklist.SetText(8, pGrid_Point, .Text)
        frmInterface.vasID.RowHeight(-1) = 12
    
'''        '바코드번호로 환자정보 불러오기
'''              SQL = "SELECT DiSTINCT CHARTNO, PATNAME, PATSEX, PATAGE,COMPANY,HOSPCODE,PATJUMIN,PATNO,COMMDATE,EXAMNO,EXAMID,IOFLAG  "
'''        SQL = SQL & vbCrLf & "  FROM PAT_RES "
'''        SQL = SQL & vbCrLf & " WHERE EXAMTYPE = '" & gPart & "' "
'''        SQL = SQL & vbCrLf & "   AND BARCODE = '" & sBarcode & "'"
'''
'''
'''        Res = GetDBSelectColumn(gLocal, SQL)
'''
'''        If Res = 1 Then
'''            SetText frmInterface.vasworklist, Trim(gReadBuf(0)), pGrid_Point, colPID    '5
'''            SetText frmInterface.vasworklist, Trim(gReadBuf(0)), pGrid_Point, colPID    '5
'''            SetText frmInterface.vasworklist, Trim(gReadBuf(1)), pGrid_Point, colPName  '6
'''            SetText frmInterface.vasworklist, Trim(gReadBuf(2)), pGrid_Point, colSex    '7
'''            SetText frmInterface.vasworklist, Trim(gReadBuf(3)), pGrid_Point, colAge    '8
'''            SetText frmInterface.vasworklist, Format(Trim(gReadBuf(8)), "####-##-##"), pGrid_Point, 2
'''
'''            SetText frmInterface.vasworklist, Trim(gReadBuf(4)), pGrid_Point, 12
'''            SetText frmInterface.vasworklist, Trim(gReadBuf(5)), pGrid_Point, 13
'''            SetText frmInterface.vasworklist, Trim(gReadBuf(6)), pGrid_Point, 14
'''            SetText frmInterface.vasworklist, Trim(gReadBuf(7)), pGrid_Point, 15
'''            SetText frmInterface.vasworklist, Trim(gReadBuf(8)), pGrid_Point, 16
'''            SetText frmInterface.vasworklist, Trim(gReadBuf(9)), pGrid_Point, 17
'''            SetText frmInterface.vasworklist, Trim(gReadBuf(10)), pGrid_Point, 18
'''            SetText frmInterface.vasworklist, Trim(gReadBuf(11)), pGrid_Point, 19
'''            frmInterface.vasworklist.RowHeight(-1) = 12
'''        End If
    
    End With
End Sub
