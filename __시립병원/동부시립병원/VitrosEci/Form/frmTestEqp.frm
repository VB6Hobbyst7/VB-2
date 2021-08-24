VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form frmTestEqp 
   Caption         =   "장비 VS 검사코드 설정"
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9945
   ScaleWidth      =   15210
   WindowState     =   2  '최대화
   Begin MSComctlLib.ImageList imlList 
      Left            =   13410
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestEqp.frx":0000
            Key             =   "TST_E"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestEqp.frx":059A
            Key             =   "TST_M"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraCmdBar 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   1.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   30
      TabIndex        =   0
      Top             =   9330
      Width           =   15120
      Begin VB.CommandButton cmdAction 
         Caption         =   "Close"
         Height          =   375
         Index           =   5
         Left            =   13710
         TabIndex        =   6
         Top             =   120
         Width           =   1245
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Refresh"
         Height          =   375
         Index           =   4
         Left            =   12390
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Save"
         Height          =   375
         Index           =   3
         Left            =   11070
         TabIndex        =   4
         Top             =   120
         Width           =   1245
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Print"
         Height          =   375
         Index           =   2
         Left            =   9750
         TabIndex        =   3
         Top             =   120
         Width           =   1245
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Del"
         Height          =   375
         Index           =   1
         Left            =   9000
         TabIndex        =   2
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Add"
         Height          =   375
         Index           =   0
         Left            =   8310
         TabIndex        =   1
         Top             =   120
         Width           =   645
      End
   End
   Begin FPSpreadADO.fpSpread spdTestListDt 
      Height          =   8835
      Left            =   180
      TabIndex        =   7
      Top             =   180
      Width           =   14775
      _Version        =   393216
      _ExtentX        =   26061
      _ExtentY        =   15584
      _StockProps     =   64
      EditModePermanent=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   12
      MaxRows         =   14
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      ScrollBarShowMax=   0   'False
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmTestEqp.frx":0B34
      ScrollBarTrack  =   1
   End
End
Attribute VB_Name = "frmTestEqp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const OBJTAG_EQP    As String = "EQP"
Private Const OBJTAG_TST    As String = "TST"
Private Const AUTO_VEFY     As String = "YES"
Private Const AUTO_VEFN     As String = "NO"

Private Const TLB_TEMP      As String = "TEMPTEABLE"
Private Const TLB_RESULT    As String = "INTERFACE003"

Private mAdoRs              As ADODB.Recordset
Private WithEvents PopUp_List As Listview
Attribute PopUp_List.VB_VarHelpID = -1

Private Sub cmdAction_Click(Index As Integer)

    Select Case Index
        Case 0: Call cmdAdded
        Case 1: Call cmdDelete
        Case 2: Call cmdPrint
        Case 3: Call cmdSave
        Case 4: Call f_subSet_EqpData(INS_CODE)
        Case 5: Call cmdClose
        Case Else
    End Select
    
End Sub

Private Sub cmdPrint()
    
    Dim strPage As String
    Dim strArea As String
    Dim strPDate As String
    
    strPage = "Page : " & Space(7) & "/p" & " of " & spdTestListDt.PrintPageCount
    strArea = ""
    strPDate = "출력일자:" & Format(Now, "yyyy년mm월dd일")
    
    With SpPrint
        .strTitle = "/fn""굴림체""/fz""20""/fb1/fi0/fu1/fk0/fs1" _
                  & "/f1/c검사항목/n"
        .strBaseDate = "/fn""굴림체""/fz""10""/fb0/fi0/fu0/fk0/fs1" _
                     & "/f1/c" & "" & "/n/n"
        .strPageCount = "/fn""굴림체"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                      & "/f1/r" & strPage & "/n"
        .strAreaName = "/fn""굴림체"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                      & "/f1/l" & strArea
        .strPrintDate = "/fn""굴림체"" /fz""9"" /fb0/fi0/fu0/fk0/fs1" _
                      & "/f1/r" & strPDate & ""
    End With

    Call Load_From(frmSpPreview)

End Sub

Private Sub Load_From(ByVal frm As Form)
    
    With frm
        .Show
        .SetFocus
    End With
    
End Sub


Private Sub cmdAdded()
    
    Dim iRow As Integer
        
    With spdTestListDt
        .Col = 1
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
    End With
    
End Sub

Private Sub cmdDelete()

    Dim iRow        As Integer
    
    With spdTestListDt
        .DeleteRows .ActiveRow, 1
        .MaxRows = .MaxRows - 1
    End With
    
End Sub

Private Sub cmdClose()
    Unload Me
End Sub

Private Sub cmdSave()
    Dim iRow        As Integer
    Dim sqlDoc  As String, sqlRet   As Integer

    On Error GoTo frmTestEqp_Add_Error
    
    With spdTestListDt
        sqlDoc = "Delete from INTERFACE002 where EQP_CD = '" & INS_CODE & "'"
        AdoCn_Jet.Execute sqlDoc, sqlRet
        
        For iRow = 1 To .MaxRows
            .Row = iRow
            .Col = 1
            If Trim(.Text) <> "" Then
                                   sqlDoc = "Update INTERFACE002"
                .Col = 2: sqlDoc = sqlDoc + "   set TESTNM_EQP = '" & Trim$(.Text) & "'," & _
                                            "       OUT_SEQ    = " & iRow & ","
                .Col = 4: sqlDoc = sqlDoc + "       TESTCD     = '" & Trim$(.Text) & "',"
                .Col = 3: sqlDoc = sqlDoc + "       TESTNM     = '" & Trim$(.Text) & "'," & _
                                            "       AUTOVERIFY = '',"
                .Col = 5:  sqlDoc = sqlDoc + "      REFLM      = '" & Trim$(.Text) & "',"
                .Col = 6:  sqlDoc = sqlDoc + "      REFHM      = '" & Trim$(.Text) & "',"
                .Col = 7:  sqlDoc = sqlDoc + "      REFLF      = '" & Trim$(.Text) & "',"
                .Col = 8:  sqlDoc = sqlDoc + "      REFHF      = '" & Trim$(.Text) & "',"
                .Col = 9:  sqlDoc = sqlDoc + "      PANICL     = '" & Trim$(.Text) & "',"
                .Col = 10: sqlDoc = sqlDoc + "      PANICH     = '" & Trim$(.Text) & "',"
                .Col = 11: sqlDoc = sqlDoc + "      DELTA      = '" & Trim$(.Text) & "',"
                .Col = 12: sqlDoc = sqlDoc + "      REMARK     = '" & Trim$(.Text) & "'"
                           sqlDoc = sqlDoc + " where EQP_CD     = '" & INS_CODE & "'"
                .Col = 1:  sqlDoc = sqlDoc + "   and TESTCD_EQP = '" & Trim$(.Text) & "'"
                         
                AdoCn_Jet.Execute sqlDoc, sqlRet
                
                If sqlRet = 0 Then
                                       sqlDoc = "Insert into INTERFACE002(" & _
                                                "            EQP_CD, TESTCD_EQP, TESTNM_EQP, OUT_SEQ, TESTNM, TESTCD, AUTOVERIFY, DELTAGBN, " & _
                                                "            REFLM, REFHM, REFLF, REFHF, PANICL, PANICH, DELTA, REMARK)" & _
                                                "    values( '" & INS_CODE & "', "
                    .Col = 1: sqlDoc = sqlDoc + "            '" & Trim$(.Text) & "',"
                    .Col = 2: sqlDoc = sqlDoc + "            '" & Trim$(.Text) & "',"
                              sqlDoc = sqlDoc + "             " & iRow & ","
                    .Col = 3: sqlDoc = sqlDoc + "            '" & Trim$(.Text) & "',"
                    .Col = 4: sqlDoc = sqlDoc + "            '" & Trim$(.Text) & "',"
                              sqlDoc = sqlDoc + "            '', '', "
                    .Col = 5:  sqlDoc = sqlDoc + "           '" & Trim$(.Text) & "',"
                    .Col = 6:  sqlDoc = sqlDoc + "           '" & Trim$(.Text) & "',"
                    .Col = 7:  sqlDoc = sqlDoc + "           '" & Trim$(.Text) & "',"
                    .Col = 8:  sqlDoc = sqlDoc + "           '" & Trim$(.Text) & "',"
                    .Col = 9:  sqlDoc = sqlDoc + "           '" & Trim$(.Text) & "',"
                    .Col = 10: sqlDoc = sqlDoc + "           '" & Trim$(.Text) & "',"
                    .Col = 11: sqlDoc = sqlDoc + "           '" & Trim$(.Text) & "',"
                    .Col = 12: sqlDoc = sqlDoc + "           '" & Trim$(.Text) & "')"
                              
                    AdoCn_Jet.Execute sqlDoc, sqlRet
                End If
            End If
        Next
    End With
    
    Call f_subSet_EqpData(INS_CODE)

    Exit Sub
frmTestEqp_Add_Error:

    Call ErrMsgProc("frmTestEqp - Private Sub cmdSave()")

End Sub

Private Sub Form_Load()
    
    Call f_subSet_EqpData(INS_CODE)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If PopUp_List Is Nothing Then Set PopUp_List = Nothing

End Sub

Private Sub Form_Resize()
    Dim i As Integer
    If ScaleHeight < 650 Then Exit Sub
    If ScaleWidth < 60 Then Exit Sub
    
    fraCmdBar.Move ScaleLeft + 30, ScaleHeight - fraCmdBar.Height - 30, ScaleWidth - 60
    For i = cmdAction.LBound To cmdAction.UBound
'        Call cmdAction(i).Move(fraCmdBar.Width - ((1300 * (cmdAction.Count - i)) + (70 * (cmdAction.UBound - i)) + 100), _
                               (fraCmdBar.Height - 360) / 2, 1300, 360)
        Call cmdAction(i).Move(fraCmdBar.Width - ((1300 * (cmdAction.Count - i)) + (70 * (cmdAction.UBound - i)) + 100), _
                               (fraCmdBar.Height - 360) / 2, 1300, 360)
    Next
    
End Sub

Private Sub f_subSet_EqpData(ByVal strEqp_Cd As String)

    Dim adoRS   As New ADODB.Recordset
    Dim sqlDoc  As String
    Dim iRow    As Integer
    
    sqlDoc = "select TESTCD_EQP, TESTNM_EQP, OUT_SEQ, TESTCD, TESTNM," & _
             "       AUTOVERIFY, REMARK, REFLM, REFHM, REFLF, REFHF, DELTA," & _
             "       DELTAGBN,   PANICL,     PANICH" & _
             "  from INTERFACE002" & _
             " where EQP_CD = '" & INS_CODE & "'" & _
             " order by OUT_SEQ, TESTCD_EQP, TESTCD"
             
    adoRS.CursorLocation = adUseClient
    adoRS.Open sqlDoc, AdoCn_Jet
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst: iRow = 0
    Do While Not adoRS.EOF
        iRow = iRow + 1
        
        With spdTestListDt
            If Trim(adoRS("TESTCD")) <> "" Then
                If iRow >= .MaxRows Then .MaxRows = .MaxRows + 1
                .Row = iRow
                .SetText 1, iRow, Trim(adoRS("TESTCD_EQP"))
                .SetText 2, iRow, Trim(adoRS("TESTNM_EQP"))
                .SetText 3, iRow, Trim(adoRS("TESTNM"))
                .SetText 4, iRow, Trim(adoRS("TESTCD"))
                .SetText 5, iRow, Trim(adoRS("REFLM"))
                .SetText 6, iRow, Trim(adoRS("REFHM"))
                .SetText 7, iRow, Trim(adoRS("REFLF"))
                .SetText 8, iRow, Trim(adoRS("REFHF"))
                .SetText 9, iRow, Trim(adoRS("PANICL"))
                .SetText 10, iRow, Trim(adoRS("PANICH"))
                .SetText 11, iRow, Trim(adoRS("DELTA"))
                .SetText 12, iRow, Trim(adoRS("REMARK"))
            End If
        End With
        
        adoRS.MoveNext
    Loop
    
    spdTestListDt.MaxRows = iRow
    adoRS.Close:    Set adoRS = Nothing
    
End Sub

Private Sub f_subClear_Form()
    
'    With spdTestList
'        .MaxRows = 1
'        .Col = 1:   .Col2 = .MaxCols
'        .Row = 1:   .Row2 = .MaxRows
'        .BlockMode = True
'        .Action = ActionClearText
'        .BlockMode = False
'        .RowHeight(-1) = 12
'    End With
    
    With spdTestListDt
        .MaxRows = 1
        .Col = 1:   .Col2 = .MaxCols
        .Row = 1:   .Row2 = .MaxRows
        .BlockMode = True
        .Action = ActionClearText
        .BlockMode = False
        .RowHeight(-1) = 12
    End With
    
'    txtTstcdEqp = ""
'    txtTstnmEqp = ""
'    txtTstcdEqpDt = ""
'    txtTstnmEqpDt = ""
'    txtTestCd = ""
    
End Sub


