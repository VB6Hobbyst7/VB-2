VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmWorkList_1 
   Caption         =   "WorkList"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   11640
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CheckBox chkAll 
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   165
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "출력"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2340
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   7860
      Visible         =   0   'False
      Width           =   1395
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   7875
      Left            =   60
      TabIndex        =   9
      Top             =   870
      Width           =   11505
      _Version        =   393216
      _ExtentX        =   20294
      _ExtentY        =   13891
      _StockProps     =   64
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   8
      MaxRows         =   100
      ScrollBars      =   2
      ShadowColor     =   15526606
      ShadowDark      =   13815180
      SpreadDesigner  =   "frmWorkList_1.frx":0000
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   735
      Left            =   60
      TabIndex        =   4
      Top             =   90
      Width           =   11505
      _Version        =   65536
      _ExtentX        =   20294
      _ExtentY        =   1296
      _StockProps     =   15
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Begin VB.ComboBox cboExam 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6210
         TabIndex        =   20
         Text            =   "검사항목"
         Top             =   195
         Width           =   1905
      End
      Begin VB.TextBox txtRingID 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5220
         TabIndex        =   19
         Text            =   "999999"
         Top             =   195
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton cmdWorkList 
         Caption         =   "전송"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   9240
         Style           =   1  '그래픽
         TabIndex        =   17
         Top             =   120
         Width           =   1035
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "닫기"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   10290
         Style           =   1  '그래픽
         TabIndex        =   16
         Top             =   120
         Width           =   1035
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "조회"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   8190
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   120
         Width           =   1035
      End
      Begin VB.ComboBox cboCha 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         Style           =   1  '단순 콤보
         TabIndex        =   11
         Top             =   195
         Width           =   705
      End
      Begin VB.CheckBox chkInOut 
         BackColor       =   &H00E0E0E0&
         Caption         =   "외래"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1500
         TabIndex        =   6
         Top             =   1020
         Value           =   1  '확인
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CheckBox chkInOut 
         BackColor       =   &H00E0E0E0&
         Caption         =   "입원"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2460
         TabIndex        =   5
         Top             =   1020
         Value           =   1  '확인
         Visible         =   0   'False
         Width           =   765
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   315
         Left            =   1410
         TabIndex        =   10
         Top             =   510
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   21364737
         CurrentDate     =   38888
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   315
         Left            =   1200
         TabIndex        =   12
         Top             =   195
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   21364737
         CurrentDate     =   38888
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "RingID"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4440
         TabIndex        =   18
         Top             =   255
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "작업일자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   255
         Width           =   990
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "차수"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3000
         TabIndex        =   14
         Top             =   255
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3000
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "외래/입원"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   7
         Top             =   1050
         Visible         =   0   'False
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmdDown 
      Height          =   495
      Left            =   840
      Picture         =   "frmWorkList_1.frx":11CF
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   8010
      Width           =   705
   End
   Begin VB.CommandButton cmdUp 
      Height          =   495
      Left            =   90
      Picture         =   "frmWorkList_1.frx":1301
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   8010
      Width           =   705
   End
End
Attribute VB_Name = "frmWorkList_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAll_Click()
    vasList.Row = -1
    vasList.Col = 1
    
    If chkAll.Value = 0 Then
        vasList.Value = 0
    ElseIf chkAll.Value = 1 Then
        vasList.Value = 1
    End If
End Sub

Private Sub cmdDown_Click()
    Dim lRow As Long
    
    lRow = vasList.ActiveRow
    
    vasList.SwapRange 1, lRow, 11, lRow, 1, lRow + 1
    vasActiveCell vasList, lRow + 1, 2
    vasList_Click 2, lRow + 1
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    Dim lsDate As String
    Dim liOrdNo As Integer
    Dim i, j, k, n
    Dim lRow, lCol As Long
    
    Dim lsWorkStation As String
    
    Dim rsBarcode As ADODB.Recordset
    Dim cmdBarcode As New ADODB.Command
    
    Dim iOrd As Integer
    
    On Error GoTo errtrap
    
    ClearSpread vasList
    
    If Trim(cboCha.Text) = "" Then
        MsgBox "차수를 입력하세요!", vbInformation + vbOKOnly, "알림"
        cboCha.SetFocus
        Exit Sub
    End If
        
    'lsWorkStation = "13"
    lsWorkStation = "79"
    'lsWorkStation = "01"
    
    Me.MousePointer = 11
    
    lsDate = dtpSDate.Value
    
    lRow = 1
            
    If Not rs Is Nothing Then Set rs = Nothing
    
    With cmdSQL
        .ActiveConnection = cn_Ser
        .CommandType = adCmdStoredProc
        .CommandText = "Interface_WL_List_SELECT_sp"
        '.Parameters.Append .CreateParameter("retval", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@i_instrumentcode", adChar, adParamInput, 11, lsWorkStation)
        .Parameters.Append .CreateParameter("@i_WorkList_Date", adChar, adParamInput, 11, lsDate)
        .Parameters.Append .CreateParameter("@i_Order_Number", adChar, adParamInput, 11, Trim(cboCha.Text))
        .Parameters.Append .CreateParameter("@i_from_seq_number", adChar, adParamInput, 11, "1")
        .Parameters.Append .CreateParameter("@i_to_seq_number", adChar, adParamInput, 11, "1000")

        Set rs = New ADODB.Recordset
        rs.CursorType = adOpenStatic
        Set rs = .Execute
    End With

    For n = 0 To cmdSQL.Parameters.Count - 1
        cmdSQL.Parameters.Delete 0
    Next n
    
    While Not rs.EOF
        If vasList.MaxRows < lRow Then
            vasList.MaxRows = lRow
        End If
        
        iOrd = -1
        With cmdSQL
            .ActiveConnection = cn
            .CommandType = adCmdText
            .CommandText = "select barcode, OrdFlag from worklist where barcode = '" & Trim(rs.Fields.Item(1).Value) & "' "
            Set rsBarcode = New ADODB.Recordset
            Set rsBarcode = .Execute
        End With
        If Not rsBarcode.EOF Then
            If Trim(CStr(rsBarcode.Fields.Item(1).Value)) = "B" Then
                iOrd = 1
            End If
            rsBarcode.Close
        End If
        
        With cmdSQL
            .ActiveConnection = cn
            .CommandType = adCmdText
            .CommandText = "select barcode from pat_res where barcode = '" & Trim(rs.Fields.Item(1).Value) & "' "
            Set rsBarcode = New ADODB.Recordset
            Set rsBarcode = .Execute
        End With
        If Not rsBarcode.EOF Then
            If Trim(CStr(rsBarcode.Fields.Item(0).Value)) = Trim(rs.Fields.Item(1).Value) Then
                iOrd = 1
            End If
            rsBarcode.Close
        End If
        
        If Not rsBarcode Is Nothing Then Set rsBarcode = Nothing
        If Not cmdBarcode Is Nothing Then Set cmdBarcode = Nothing

        'If iOrd <> 1 Then
            vasList.Row = lRow
            vasList.Col = 2
            If IsNull(rs.Fields.Item(0).Value) Then
                vasList.Text = ""
            Else
                vasList.Text = Trim(cboCha.Text) & "-" & SetChar(Trim(CStr(rs.Fields.Item(0).Value)), 2, 1, " ")
            End If
            
            For lCol = 1 To rs.Fields.Count - 3
                vasList.Row = lRow
                vasList.Col = lCol + 2
                If IsNull(rs.Fields.Item(lCol).Value) Then
                    vasList.Text = ""
                Else
                    vasList.Text = Trim(CStr(rs.Fields.Item(lCol).Value))
                End If
            Next lCol
            vasList.Row = lRow
            vasList.Col = 1
            If iOrd = 1 Then
                vasList.Value = 0
            Else
                vasList.Value = 1
            End If
            
            lRow = lRow + 1
        'End If
        
        rs.MoveNext
    Wend
    
    vasList.MaxRows = vasList.DataRowCnt
    
    Me.MousePointer = 0
    
    vasList.RowHeight(-1) = 12
    
    Exit Sub
    
errtrap:
    Me.MousePointer = 0
    
    If Not rs Is Nothing Then Set rs = Nothing
    If Not cmdSQL Is Nothing Then Set cmdSQL = Nothing
    
    If Not rsBarcode Is Nothing Then Set rsBarcode = Nothing
    If Not cmdBarcode Is Nothing Then Set cmdBarcode = Nothing
    
    'MsgBox Err.Number & " : " & Err.Description

    Exit Sub
End Sub

Private Sub cmdUp_Click()
    Dim lRow As Long
    
    lRow = vasList.ActiveRow
    
    vasList.SwapRange 1, lRow, 11, lRow, 1, lRow - 1
    vasActiveCell vasList, lRow - 1, 2
    vasList_Click 2, lRow - 1
End Sub

Private Sub cmdWorkList_Click()
    Dim lRow As Long
    Dim lCol As Long
    Dim lDestRow As Long
    
    Dim mExam As Variant
    Dim lsID As String
    Dim i, j, k, z As Long
    Dim lsEquipCode As String
    Dim lsOrder As String
    Dim lsTest As String
    Dim lsOrdDate As String
    Dim liTube As Integer
    Dim liRing As Double
    Dim liStart As Integer
    
    Dim lsSpcID As String
    Dim sExamDate As String
    Dim lsSeqNo As String
    Dim lsExamName As String
    Dim lsPreEquipCode As String
    
    lDestRow = frmInterface.vasList.DataRowCnt + 1
    If lDestRow > frmInterface.vasList.MaxRows Then
        frmInterface.vasList.MaxRows = lDestRow
    End If

'    If frmInterface.vasID.DataRowCnt = 1 Then
'        lDestRow = lDestRow - 1
'    End If
    
    If Not IsNumeric(txtRingID) Then
        MsgBox "RingID를 입력하세요"
        txtRingID = ""
        txtRingID.SetFocus
        Exit Sub
    End If
    
    lsOrdDate = Format(GetDateFull, "mm.dd.yyyy hh:nn:ss")
    sExamDate = GetDateFull
    
    lsOrder = ""
    liTube = 0
    liRing = CDbl(txtRingID)
    liStart = 1
    For lRow = 1 To vasList.DataRowCnt
        vasList.Row = lRow
        vasList.Col = 1
        If vasList.Value = 1 Then
'            For lCol = 2 To 7
'                SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, lCol + 2
'            Next lCol
            lsID = Trim(GetText(vasList, lRow, 3))
            lsSpcID = Trim(GetText(vasList, lRow, 2))
            i = InStr(1, lsSpcID, "-")
            If i > 0 Then
                lsSpcID = SetSpace(Trim(Left(lsSpcID, i - 1)), 2) & SetSpace(Trim(Mid(lsSpcID, i + 1)), 2)
                lsSpcID = Format(dtpSDate.Value, "yymmdd") & lsSpcID
            End If
                
            lsPreEquipCode = ""
            mExam = Get_OrderBody(lsID)
            k = 0
            'lsOrd = ""
            If Not IsNull(mExam) Then
                lsTest = ""
                
                For i = LBound(mExam, 2) To UBound(mExam, 2)
                    'SetText vasExam, mExam(3, i), lRow, 1
                    'SetText vasExam, mExam(4, i), lRow, 2
                    
                    gReadBuf(0) = ""
                    gReadBuf(1) = ""
                    
                    SQL = "Select ExamCode, EquipCode, ExamName, SeqNo from EquipExam " & vbCrLf & _
                          "where Equip = '" & gEquip & "' " & vbCrLf & _
                          "  and ExamCode = '" & Trim(mExam(3, i)) & "' " & vbCrLf & _
                          "  and UseFlag = 1 "
                    res = db_select_Col(gLocal, SQL)
                    If Trim(gReadBuf(0)) = Trim(mExam(3, i)) Then
                        If lsPreEquipCode = Trim(gReadBuf(1)) Then
                            gReadBuf(1) = ""
                        End If
                        
                        If lsOrder = "" And cboExam.ListIndex < 1 And Trim(gReadBuf(1)) <> "" Then
                            cboExam.Text = Trim(gReadBuf(1))
                        End If
                        
                        If Trim(cboExam.Text) = Trim(gReadBuf(1)) And Trim(gReadBuf(1)) <> "" Then
                            If lsTest = "" Then
                                liTube = liTube + 1
                                If liTube > 12 Then
                                    liRing = liRing + 1
                                    liTube = 1
                                      
                                    liStart = 1
                                End If
                                                    
    '                            If liStart = 1 Then
    '                                lsOrder = lsOrder & "01 " & SetSpace(CStr(liRing), 6, 1) & End_Char 'A-Ring ID (6자리)
    '                                lsOrder = lsOrder & "02 " & lsOrdDate & End_Char 'Order Date/Time
    '                                lsOrder = lsOrder & "03 3" & End_Char 'Order Run Mode
    '                            End If
    '                            liStart = 0
                            End If
                            lsEquipCode = Trim(gReadBuf(1))
                            'lsTest = lsTest & "07 " & SetSpace(lsEquipCode, 3) & " 1" & End_Char 'TestID
                            lsTest = lsTest & "07 " & SetSpace(lsEquipCode, 3) & End_Char 'TestID
                            lsExamName = Trim(gReadBuf(2))
                            lsSeqNo = Trim(gReadBuf(3))
                            'lsTest = lsTest & "08 1" & End_Char   'Test Type
                            
                            lsPreEquipCode = lsEquipCode
                            'sCnt = ""
                            SQL = "Delete FROM pat_res " & vbCrLf & _
                                  "WHERE examdate = '" & Format(CDate(frmInterface.dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
                                  "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                                  "  AND equipcode = '" & lsEquipCode & "'" & vbCrLf & _
                                  "  AND barcode = '" & lsID & "' "
                            
                            res = SendQuery(gLocal, SQL)
                            If res = -1 Then
                                SaveQuery SQL
                                'Exit
                            End If
                            
                            SQL = "INSERT INTO pat_res (examdate, equipno, " & _
                                    "barcode, examtype, receno, " & _
                                    "pid, pname, pjumin, page, psex, " & _
                                    "resdate, seqno, diskno, posno, " & _
                                    "equipcode, examcode, " & _
                                    "result, sendflag, examname, " & _
                                    "refflag,panicflag, deltaflag, unit, refvalue, panicvalue, examuid ) " & vbCrLf & _
                                  "VALUES ('" & Format(CDate(frmInterface.dtpExamDate.Value), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
                                  "'" & lsID & "','', '" & Trim(GetText(vasList, lRow, 2)) & "', " & _
                                  "'" & Trim(GetText(vasList, lRow, 4)) & "', '" & Trim(GetText(vasList, lRow, 5)) & "', '', 0, '', " & _
                                  "'" & sExamDate & "', '" & lsSeqNo & "', '" & liRing & "', '" & liTube & "', " & vbCrLf & _
                                  "'" & lsEquipCode & "', '" & Trim(mExam(3, i)) & "', " & _
                                  "'', 'A', '" & lsExamName & "', " & vbCrLf & _
                                  "'', '', '', '', " & _
                                  "'', '', '" & lsSpcID & "' ) "
                            res = SendQuery(gLocal, SQL)
                            If res = -1 Then
                                SaveQuery SQL
                                'Exit Function
                            End If
                        End If
                    End If
                Next i
                         
                If lsTest <> "" Then
'                    liTube = liTube + 1
'                    If liTube > 12 Then
'                        liRing = liRing + 1
'                        liTube = 1
'
'                        liStart = 1
'                    End If
                                        
                    If liStart = 1 Then
                        lsOrder = lsOrder & "01 " & SetSpace(CStr(liRing), 6, 1) & End_Char 'A-Ring ID (6자리)
                        lsOrder = lsOrder & "02 " & lsOrdDate & End_Char 'Order Date/Time
                        lsOrder = lsOrder & "03 3" & End_Char 'Order Run Mode
                    End If
                    liStart = 0
                    lsOrder = lsOrder & "04 " & SetSpace(CStr(liTube), 2, 1) & End_Char 'A-tub Position (2자리)
                    lsOrder = lsOrder & "05 2" & End_Char 'Order Type
                    lsOrder = lsOrder & "06 " & lsSpcID & End_Char 'Specimen Infomation(10자리)
                    lsOrder = lsOrder & lsTest
                    
                    SetText vasList, "오더", lRow, gResCol
                    SetBackColor vasList, lRow, lRow, 1, 1, 255, 250, 205
                    
                        
                    frmInterface.vasList.SetText 2, lDestRow, Trim(GetText(vasList, lRow, 3))   '바코드
                    frmInterface.vasList.SetText 3, lDestRow, Trim(GetText(vasList, lRow, 4))   '등록번호
                    frmInterface.vasList.SetText 4, lDestRow, Trim(GetText(vasList, lRow, 5))   '환자이름
                    frmInterface.vasList.SetText 5, lDestRow, Trim(GetText(vasList, lRow, 2))   '차수
                    frmInterface.vasList.SetText 6, lDestRow, ""    'Rack
                    frmInterface.vasList.SetText 7, lDestRow, ""    'Pos
                    frmInterface.vasList.SetText 8, lDestRow, "오더"    '상태
                    
                    vasList.SetText 7, lRow, "오더"
                    vasList.SetText 8, lRow, lsSpcID
                    
                    vasList.Row = lRow
                    vasList.Col = 1
                    vasList.Value = 0
                    vasList.BackColor = RGB(255, 255, 255)
                    
                    'lDestRow = lDestRow + 145
                    lDestRow = lDestRow + 1
                    
                    If lDestRow > frmInterface.vasList.MaxRows Then
                        frmInterface.vasList.MaxRows = lDestRow
                    End If
                Else
                    SQL = "Delete FROM pat_res " & vbCrLf & _
                          "WHERE examdate = '" & Format(CDate(frmInterface.dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
                          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                          "  AND equipcode = '" & lsEquipCode & "'" & vbCrLf & _
                          "  AND barcode = '" & lsID & "' "
                    
                    res = SendQuery(gLocal, SQL)
                    If res = -1 Then
                        SaveQuery SQL
                        'Exit
                    End If
                    
                    SQL = "INSERT INTO pat_res (examdate, equipno, " & _
                            "barcode, examtype, receno, " & _
                            "pid, pname, pjumin, page, psex, " & _
                            "resdate, seqno, diskno, posno, " & _
                            "equipcode, examcode, " & _
                            "result, sendflag, examname, " & _
                            "refflag,panicflag, deltaflag, unit, refvalue, panicvalue, examuid ) " & vbCrLf & _
                          "VALUES ('" & Format(CDate(frmInterface.dtpExamDate.Value), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
                          "'" & lsID & "','', '" & Trim(GetText(vasList, lRow, 2)) & "', " & _
                          "'" & Trim(GetText(vasList, lRow, 4)) & "', '" & Trim(GetText(vasList, lRow, 5)) & "', '', 0, '', " & _
                          "'" & sExamDate & "', '" & lsSeqNo & "', '" & liRing & "', '" & liTube & "', " & vbCrLf & _
                          "'" & lsEquipCode & "', '', " & _
                          "'', 'A', '" & lsExamName & "', " & vbCrLf & _
                          "'', '', '', '', " & _
                          "'', '', '" & lsSpcID & "' ) "
                    res = SendQuery(gLocal, SQL)
                    If res = -1 Then
                        SaveQuery SQL
                        'Exit Function
                    End If
                    
                    frmInterface.vasList.SetText 2, lDestRow, Trim(GetText(vasList, lRow, 3))   '바코드
                    frmInterface.vasList.SetText 3, lDestRow, Trim(GetText(vasList, lRow, 4))   '등록번호
                    frmInterface.vasList.SetText 4, lDestRow, Trim(GetText(vasList, lRow, 5))   '환자이름
                    frmInterface.vasList.SetText 5, lDestRow, Trim(GetText(vasList, lRow, 2))   '차수
                    frmInterface.vasList.SetText 6, lDestRow, ""    'Rack
                    frmInterface.vasList.SetText 7, lDestRow, ""    'Pos
                    frmInterface.vasList.SetText 8, lDestRow, "없음"    '상태
                    
                    'lDestRow = lDestRow + 145
                    lDestRow = lDestRow + 1
                    
                    If lDestRow > frmInterface.vasList.MaxRows Then
                        frmInterface.vasList.MaxRows = lDestRow
                    End If
                    
                    vasList.SetText 7, lRow, "없음"
                    vasList.Row = lRow
                    vasList.Col = 1
                    vasList.Value = 1
                    vasList.BackColor = RGB(255, 0, 0)
                End If
            Else
                
                vasList.SetText 7, lRow, "오류"
                vasList.Row = lRow
                vasList.Col = 1
                vasList.Value = 1
                vasList.BackColor = RGB(255, 0, 0)
            End If
            
        End If
    Next lRow
    
    frmInterface.vasList.RowHeight(-1) = 12
        
    'chkAll.Value = 0
    
    If lsOrder <> "" Then
        lsOrder = lsOrder & "20 19   +6.00E+01  +5.40E+02  +5.50E+03  +2.20E+04" & End_Char
        'Amplicor_Order_Entry lsOrder
    End If
    
    gRow = 0
    
    Unload Me
End Sub

Private Sub Form_Load()
    '조회일자
    dtpSDate.Value = Trim(frmInterface.dtpExamDate.Value)
    dtpEDate.Value = Trim(dtpSDate.Value)
    
    ClearSpread vasList
    
    cboExam.Clear
    
    SQL = "Select EquipCode from EquipExam where Equip = '" & gEquip & "' and UseFlag = 1 "
    res = db_select_Combo(gLocal, SQL, cboExam)
    cboExam.AddItem "검사선택", 0
    cboExam.ListIndex = 0
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 0 Or Row > vasList.DataRowCnt Then
        cmdUp.Enabled = False
        cmdDown.Enabled = False
    End If
    
    If Row = 1 Then
        cmdUp.Enabled = False
        cmdDown.Enabled = True
    ElseIf Row = vasList.DataRowCnt Then
        cmdUp.Enabled = True
        cmdDown.Enabled = False
    Else
        cmdUp.Enabled = True
        cmdDown.Enabled = True
    End If
    If Col = 4 And Row = 0 Then
        vasSort vasList, 4
    End If
    If Col = 2 And Row = 0 Then
        vasSort vasList, 2
    End If
End Sub
