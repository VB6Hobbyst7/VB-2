VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "spr32x60.ocx"
Begin VB.Form frmWorkList 
   Caption         =   "WorkList"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12810
   LinkTopic       =   "Form1"
   ScaleHeight     =   8835
   ScaleWidth      =   12810
   StartUpPosition =   1  '소유자 가운데
   Begin FPSpread.vaSpread vasTemp1 
      Height          =   915
      Left            =   1470
      TabIndex        =   24
      Top             =   5460
      Visible         =   0   'False
      Width           =   2805
      _Version        =   393216
      _ExtentX        =   4948
      _ExtentY        =   1614
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
      SpreadDesigner  =   "frmWorkList.frx":0000
   End
   Begin FPSpread.vaSpread vasTemp 
      Height          =   1785
      Left            =   360
      TabIndex        =   19
      Top             =   7170
      Visible         =   0   'False
      Width           =   10125
      _Version        =   393216
      _ExtentX        =   17859
      _ExtentY        =   3149
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
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      ShadowColor     =   15526606
      ShadowDark      =   13815180
      SpreadDesigner  =   "frmWorkList.frx":0205
   End
   Begin VB.CheckBox chkAll 
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   630
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
      Width           =   12645
      _Version        =   393216
      _ExtentX        =   22304
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
      MaxCols         =   9
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      ShadowColor     =   15526606
      ShadowDark      =   13815180
      SpreadDesigner  =   "frmWorkList.frx":2631
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   765
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   12645
      _Version        =   65536
      _ExtentX        =   22304
      _ExtentY        =   1349
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
      Begin Versant440_국립암센터.MDButton cmdExit 
         Height          =   525
         Left            =   11520
         TabIndex        =   23
         Top             =   120
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "닫기"
      End
      Begin Versant440_국립암센터.MDButton cmdWorkList 
         Height          =   525
         Left            =   9660
         TabIndex        =   22
         Top             =   120
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "작성"
      End
      Begin Versant440_국립암센터.MDButton cmdSearch1 
         Height          =   525
         Left            =   8730
         TabIndex        =   21
         Top             =   120
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "조회"
      End
      Begin Versant440_국립암센터.MDButton cmdRingIDSet 
         Height          =   345
         Left            =   5070
         TabIndex        =   20
         Top             =   1470
         Visible         =   0   'False
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Set"
      End
      Begin VB.ComboBox cboExam 
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
         Left            =   6240
         TabIndex        =   18
         Text            =   "전체선택"
         Top             =   210
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
         Left            =   4350
         TabIndex        =   17
         Text            =   "1"
         Top             =   1485
         Visible         =   0   'False
         Width           =   735
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
         Left            =   3180
         Style           =   1  '그래픽
         TabIndex        =   13
         Top             =   1440
         Visible         =   0   'False
         Width           =   1005
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
         ItemData        =   "frmWorkList.frx":4ABB
         Left            =   5820
         List            =   "frmWorkList.frx":4ABD
         TabIndex        =   11
         Top             =   1245
         Visible         =   0   'False
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
         Left            =   1530
         TabIndex        =   6
         Top             =   1410
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
         Left            =   2490
         TabIndex        =   5
         Top             =   1410
         Value           =   1  '확인
         Visible         =   0   'False
         Width           =   765
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   315
         Left            =   3240
         TabIndex        =   10
         Top             =   210
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   93716481
         CurrentDate     =   38888
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   315
         Left            =   1350
         TabIndex        =   12
         Top             =   210
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   93716481
         CurrentDate     =   38888
      End
      Begin Versant440_국립암센터.MDButton MDButton1 
         Height          =   525
         Left            =   10590
         TabIndex        =   25
         Top             =   120
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "출력"
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "검사항목"
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
         Left            =   5130
         TabIndex        =   26
         Top             =   270
         Width           =   990
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "PAT/ID"
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
         Left            =   3570
         TabIndex        =   16
         Top             =   1545
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
         Left            =   300
         TabIndex        =   15
         Top             =   285
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
         Left            =   5400
         TabIndex        =   14
         Top             =   1305
         Visible         =   0   'False
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
         Top             =   270
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
         Left            =   330
         TabIndex        =   7
         Top             =   1440
         Visible         =   0   'False
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmdDown 
      Height          =   495
      Left            =   840
      Picture         =   "frmWorkList.frx":4ABF
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   8010
      Width           =   705
   End
   Begin VB.CommandButton cmdUp 
      Height          =   495
      Left            =   90
      Picture         =   "frmWorkList.frx":4BF1
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   8010
      Width           =   705
   End
End
Attribute VB_Name = "frmWorkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iRow1, iRow2 As Long

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

Private Sub cmdRingIDSet_Click()
    Dim i, j As Long
    
    If iRow1 < 1 Or iRow2 < 1 Then Exit Sub
        
    j = 0
    For i = iRow1 To iRow2
        j = j + 1
        With vasList
            '.SetText 7, i, Trim(txtRingID)
            .SetText 8, i, Trim(txtRingID)
            
            .Row = i
            .Col = 1
            .Value = 1
            
            If IsNumeric(txtRingID) = True Then
                txtRingID = txtRingID + 1
            End If
        End With
    Next i
End Sub

Private Sub cmdSearch_Click()
    Dim mExam As Variant
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
        
    lsWorkStation = gWkCode
    
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

Private Sub cmdSearch1_Click()
    Dim mExam As Variant
    Dim lsDate As String
    Dim lsDate1 As String
    Dim liOrdNo As Integer
    Dim i, j, k, n, X
    Dim lRow, lCol As Long
    Dim sPRes As String
    Dim lsExamCode As String
    Dim lsEquipCode As String
    Dim lsSeqNo As String
    Dim lsExamName As String
    Dim lsAllExam As String
    
    Dim lsEquipSelect As String
    
    Dim lsWorkStation As String
    
'    Dim rsBarcode As ADODB.Recordset
'    Dim cmdBarcode As New ADODB.Command
    
    Dim iOrd As Integer
    Dim lsExamCnt As Integer
    Dim lsCntCode As String
    
    On Error GoTo errtrap
    
    ClearSpread vasTemp
    ClearSpread vasList
    
'    If Trim(cboCha.Text) = "" Then
'        MsgBox "차수를 입력하세요!", vbInformation + vbOKOnly, "알림"
'        cboCha.SetFocus
'        Exit Sub
'    End If
        
    lsWorkStation = gWkCode
    
    Me.MousePointer = 11
    
    lsDate = Format(dtpSDate.Value, "yyyymmdd")
    lsDate1 = Format(dtpEDate.Value, "yyyymmdd")
    
    If cboExam.Text = "전체선택" Then
        lsAllExam = "ALL"
    Else
        i = InStr(1, cboExam.Text, " ")
        If i > 0 Then
            lsAllExam = Trim(Mid(cboExam.Text, 1, i - 1))
        End If
    End If
    
    
    
    lRow = 1
    
    sPRes = Online_TLA(gXml_S13, lsDate, lsDate1)

    
    For X = 0 To giIndex
        SQL = "Select ExamCode, EquipCode, ExamName, SeqNo from EquipExam " & vbCrLf & _
              "where Equip = '" & gEquip & "' " & vbCrLf & _
              "  and ExamCode = '" & Trim(gS13_WorkList(X).TEST_CODE) & "' " & vbCrLf & _
              "  and UseFlag = 1 "
        res = db_select_Col(gLocal, SQL)
        
        If res > 0 Then
            lsExamCode = Trim(gReadBuf(0))
            lsEquipCode = Trim(gReadBuf(1))
            lsExamName = Trim(gReadBuf(2))
            lsSeqNo = Trim(gReadBuf(3))
        
            SQL = "select posno FROM pat_res " & vbCrLf & _
                  "WHERE recedate = '" & Trim(gS13_WorkList(X).ACPT_DATE) & "' " & vbCrLf & _
                  "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                  "  AND examcode = '" & gS13_WorkList(X).TEST_CODE & "'" & vbCrLf & _
                  "  AND barcode = '" & Trim(gS13_WorkList(X).SPC_NO) & "' "
            res = db_select_Col(gLocal, SQL)
            If res = -1 Then
                SaveQuery SQL
                'Exit Function
            End If
            
            If res = 0 Then
            
                SQL = "select max(posno) from pat_res " & vbCrLf & _
                      "where examdate ='" & Format(Date, "yyyymmdd") & "' and equipcode = '" & lsEquipCode & "'"
                res = db_select_Col(gLocal, SQL)
                If Trim(gReadBuf(0)) = "" Then
                    lsExamCnt = 1
                Else
                    lsExamCnt = CInt(gReadBuf(0)) + 1
                End If
            
                SQL = "INSERT INTO pat_res (examdate, equipno, " & _
                        "barcode, examtype, receno, " & _
                        "pid, pname, pjumin, page, psex, " & _
                        "recedate, seqno, diskno, posno, " & _
                        "equipcode, examcode, " & _
                        "result, sendflag, examname, " & _
                        "refflag,panicflag, deltaflag, unit, refvalue, panicvalue ) " & vbCrLf & _
                      "VALUES ('" & Format(Date, "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
                      "'" & Trim(gS13_WorkList(X).SPC_NO) & "','', '" & Trim(gS13_WorkList(X).ACPT_NO) & "', " & _
                      "'" & Trim(gS13_WorkList(X).PT_NO) & "', '" & Trim(gS13_WorkList(X).PT_NM) & "', '', 0, '', " & _
                      "'" & Trim(gS13_WorkList(X).ACPT_DATE) & "', '', '', " & lsExamCnt & ", " & vbCrLf & _
                      "'" & lsEquipCode & "', '" & lsExamCode & "', " & _
                      "'', 'O', '" & lsExamName & "', " & vbCrLf & _
                      "'', '', '', '', " & _
                      "'', '' ) "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    'Exit Function
                End If
            End If
        End If
    Next
    
    SQL = "select '', receno, barcode, pid, pname, posno, examcode, examname, equipcode from pat_res " & vbCrLf & _
          "where recedate between '" & lsDate & "' and '" & lsDate1 & "' and sendflag <> 'C' "
        If lsAllExam = "ALL" Then
        Else
            SQL = SQL & "and equipcode = '" & lsAllExam & "' "
        End If

    SQL = SQL & " group by examname, examcode, receno, barcode, posno, pid, pname, diskno, equipcode"
    res = db_select_Vas(gLocal, SQL, vasList)
    
    lsExamCnt = 1
    lsCntCode = ""
'    For i = 1 To vasList.DataRowCnt
'        If lsCntCode = Trim(GetText(vasList, i, 7)) Then
'            SetText vasList, lsExamCnt, i, 6
'            lsExamCnt = lsExamCnt + 1
'        Else
'            lsExamCnt = 1
'            lsCntCode = Trim(GetText(vasList, i, 7))
'            SetText vasList, lsExamCnt, i, 6
'            lsExamCnt = lsExamCnt + 1
'        End If
'    Next

    vasList.MaxRows = vasList.DataRowCnt
    
    Me.MousePointer = 0
    
    vasList.RowHeight(-1) = 12
    
    Exit Sub
    
errtrap:
    Me.MousePointer = 0

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
    Dim lsExamCode As String
    
    Dim lsPreEquipCode As String
    Dim X As Integer
    
    
    lsOrdDate = Format(GetDateFull, "mm.dd.yyyy hh:nn:ss")
    sExamDate = GetDateFull
    
    lsOrder = ""
    liTube = 0
    
    liStart = 1
    For lRow = 1 To vasList.DataRowCnt
        vasList.Row = lRow
        vasList.Col = 1
        If vasList.Value = 1 Then
            lsID = Trim(GetText(vasList, lRow, 3))
            lsExamCode = Trim(GetText(vasList, lRow, 7))
            
            lsPreEquipCode = ""
            ClearSpread vasTemp1
            SQL = "select  equipcode from pat_res where " & vbCrLf & _
                  "barcode = '" & lsID & "' and examcode = '" & lsExamCode & "' and sendflag <> 'C'"
            res = db_select_Col(gLocal, SQL)
            
            k = 0
            
            If res > 0 Then
                lsPreEquipCode = gReadBuf(0)
                lsEquipCode = lsPreEquipCode
                SQL = "update pat_res set " & vbCrLf & _
                      "sendflag = 'A', resdate = '" & sExamDate & "', examdate = '" & Format(Date, "yyyymmdd") & "' " & vbCrLf & _
                      "WHERE equipno = '" & gEquip & "' " & vbCrLf & _
                      "  AND equipcode = '" & lsEquipCode & "'" & vbCrLf & _
                      "  AND barcode = '" & lsID & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    'Exit Function
                End If
                lsTest = ""
               
'                If lsPreEquipCode <> "" Then
'                    SetBackColor vasList, lRow, lRow, 1, 1, 255, 250, 205
'
'
'                    frmInterface.vasList.SetText 2, lDestRow, Trim(GetText(vasList, lRow, 3))   '바코드
'                    frmInterface.vasList.SetText 3, lDestRow, Trim(GetText(vasList, lRow, 4))   '등록번호
'                    frmInterface.vasList.SetText 4, lDestRow, Trim(GetText(vasList, lRow, 5))   '환자이름
'                    frmInterface.vasList.SetText 5, lDestRow, Trim(GetText(vasList, lRow, 2))   '차수
'                    frmInterface.vasList.SetText 6, lDestRow, Trim(GetText(vasList, lRow, 7))    'Rack
'                    frmInterface.vasList.SetText 7, lDestRow, Trim(GetText(vasList, lRow, 8))    'Pos
'                    frmInterface.vasList.SetText 8, lDestRow, "오더"    '상태
'
'                    vasList.SetText 7, lRow, "오더"
'                    vasList.SetText 8, lRow, lsSpcID
'
'                    vasList.Row = lRow
'                    vasList.Col = 1
'                    vasList.Value = 0
'                    vasList.BackColor = RGB(255, 255, 255)
'
'
'                    lDestRow = lDestRow + 1
'
'                    If lDestRow > frmInterface.vasList.MaxRows Then
'                        frmInterface.vasList.MaxRows = lDestRow
'                    End If
'                Else
'
'                End If
            Else
                
                vasList.SetText 7, lRow, "오류"
                vasList.Row = lRow
                vasList.Col = 1
                vasList.Value = 1
                vasList.BackColor = RGB(255, 0, 0)
            End If
            
            
            
        End If
            
    Next lRow
    
'    frmInterface.vasList.RowHeight(-1) = 12
    
    gRow = 0
    
    gRCnt = 0
    
    frmInterface.cboExam.Text = cboExam.Text
    frmInterface.cmdSch_Click
    
    
'    Unload Me
End Sub

Private Sub dtpSDate_Change()
    Dim lsDate As String
    Dim lsWorkStation As String
    Dim n As Integer

    On Error GoTo errtrap

    cboCha.Clear

    'lsWorkStation = gInsCode
    lsWorkStation = gWkCode

    Me.MousePointer = 11

    lsDate = dtpSDate.Value

    If Not rs Is Nothing Then Set rs = Nothing

    With cmdSQL
        .ActiveConnection = cn_Ser
        .CommandType = adCmdStoredProc
        .CommandText = "WL_Order_Number_SELECT_sp"
        '.Parameters.Append .CreateParameter("retval", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@i_instrumentcode", adChar, adParamInput, 11, lsWorkStation)
        .Parameters.Append .CreateParameter("@i_WorkList_Date", adChar, adParamInput, 11, lsDate)

        Set rs = New ADODB.Recordset
        rs.CursorType = adOpenStatic
        Set rs = .Execute
    End With

    For n = 0 To cmdSQL.Parameters.Count - 1
        cmdSQL.Parameters.Delete 0
    Next n

    While Not rs.EOF

        If Not IsNull(rs.Fields.Item(0).Value) Then
            cboCha.AddItem Trim(rs.Fields.Item(0).Value)
        End If

        rs.MoveNext
    Wend

    Me.MousePointer = 0

    If cboCha.ListCount > 0 Then
        cboCha.ListIndex = cboCha.ListCount - 1
    End If

    Exit Sub

errtrap:
    Me.MousePointer = 0

    If Not rs Is Nothing Then Set rs = Nothing
    If Not cmdSQL Is Nothing Then Set cmdSQL = Nothing

    'MsgBox Err.Number & " : " & Err.Description

    Exit Sub

End Sub

Private Sub Form_Load()
    '조회일자
    dtpSDate.Value = Trim(frmInterface.dtpExamDate.Value)
    dtpEDate.Value = Trim(dtpSDate.Value)
    
    ClearSpread vasList
    
    cboExam.Clear
    
    SQL = "Select EquipCode, ExamName from EquipExam where Equip = '" & gEquip & "' and UseFlag = 1 "
    res = db_select_Combo_2(gLocal, SQL, cboExam)
    cboExam.AddItem "전체선택", 0
    cboExam.ListIndex = 0
    
    dtpSDate_Change
End Sub

Private Sub vasList_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    If BlockRow > BlockRow2 Then
        iRow1 = BlockRow2
        iRow2 = BlockRow
    Else
        iRow1 = BlockRow
        iRow2 = BlockRow2
    End If
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

Private Sub vasList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long
    
    If KeyCode = vbKeyDelete Then
        lRow = vasList.ActiveRow
        If lRow < 1 Or lRow > vasList.DataRowCnt Then Exit Sub
        
        If MsgBox("선택한 검체를 작업내용에서 삭제하시겠습니까?", vbOKCancel + vbDefaultButton2, "알림") = vbCancel Then Exit Sub
        
        DeleteRow vasList, lRow, lRow
    End If
End Sub
