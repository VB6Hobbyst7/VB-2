VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm343Template 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   10935
   ControlBox      =   0   'False
   Icon            =   "Lis343.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00F4F0F2&
      Caption         =   "엑셀받기(&E)"
      Height          =   510
      Left            =   3825
      Style           =   1  '그래픽
      TabIndex        =   23
      Top             =   7875
      Visible         =   0   'False
      Width           =   1320
   End
   Begin FPSpread.vaSpread tblExcel 
      Height          =   750
      Left            =   5910
      TabIndex        =   22
      Top             =   180
      Visible         =   0   'False
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
      _ExtentY        =   1323
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
      MaxCols         =   2
      SpreadDesigner  =   "Lis343.frx":038A
   End
   Begin VB.CommandButton cmdBrowse 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Browse"
      Height          =   510
      Left            =   300
      Style           =   1  '그래픽
      TabIndex        =   5
      Tag             =   "25612"
      Top             =   8010
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   6810
      Left            =   3105
      TabIndex        =   6
      Top             =   960
      Width           =   7380
      Begin VB.CommandButton cmdPopupList 
         BackColor       =   &H00DEDBDD&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2205
         MousePointer    =   14  '화살표와 물음표
         Picture         =   "Lis343.frx":1AFA
         Style           =   1  '그래픽
         TabIndex        =   15
         Top             =   525
         Width           =   300
      End
      Begin VB.TextBox txtVal 
         BackColor       =   &H00F1F5F4&
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   2
         Top             =   1605
         Width           =   6930
      End
      Begin VB.TextBox txtVal 
         BackColor       =   &H00F1F5F4&
         Height          =   1845
         Index           =   1
         Left            =   255
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   2460
         Width           =   7020
      End
      Begin VB.TextBox txtVal 
         BackColor       =   &H00F1F5F4&
         Height          =   2070
         Index           =   2
         Left            =   270
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   4650
         Width           =   6960
      End
      Begin VB.TextBox txtSubKey 
         BackColor       =   &H00F7FFF7&
         Height          =   330
         Left            =   270
         TabIndex        =   1
         Top             =   525
         Width           =   1920
      End
      Begin VB.CheckBox chkKeyLock 
         BackColor       =   &H00DBE6E6&
         Caption         =   "Index Key Lock Mode"
         Height          =   240
         Left            =   5100
         TabIndex        =   7
         Top             =   585
         Value           =   1  '확인
         Width           =   2175
      End
      Begin MedControls1.LisLabel lblTestnm 
         Height          =   300
         Left            =   2520
         TabIndex        =   16
         Top             =   540
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   529
         BackColor       =   13752531
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   ""
         LeftGab         =   100
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Data 1"
         Height          =   180
         Index           =   1
         Left            =   285
         TabIndex        =   11
         Top             =   1350
         Width           =   525
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Data 2"
         Height          =   180
         Index           =   2
         Left            =   285
         TabIndex        =   10
         Top             =   2175
         Width           =   525
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Data 3"
         Height          =   180
         Index           =   3
         Left            =   300
         TabIndex        =   9
         Top             =   4380
         Width           =   525
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   105
         X2              =   7320
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000011&
         X1              =   105
         X2              =   7320
         Y1              =   1095
         Y2              =   1095
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "Index"
         Height          =   180
         Index           =   0
         Left            =   285
         TabIndex        =   8
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.ListBox lstSubKey 
      BackColor       =   &H00F7FFF7&
      Height          =   6360
      Left            =   255
      TabIndex        =   0
      Top             =   1455
      Width           =   2715
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00EBF3ED&
      Height          =   690
      Left            =   3105
      ScaleHeight     =   630
      ScaleWidth      =   7320
      TabIndex        =   17
      Top             =   7800
      Width           =   7380
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00F4F0F2&
         Caption         =   "화면지움(&C)"
         Height          =   510
         Left            =   4665
         Style           =   1  '그래픽
         TabIndex        =   21
         Tag             =   "25612"
         Top             =   45
         Width           =   1320
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00F4F0F2&
         Caption         =   "저장(&S)"
         Height          =   510
         Left            =   2025
         Style           =   1  '그래픽
         TabIndex        =   20
         Tag             =   "25612"
         Top             =   45
         Width           =   1320
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00F4F0F2&
         Caption         =   "삭제(&D)"
         Height          =   510
         Left            =   3345
         Style           =   1  '그래픽
         TabIndex        =   19
         Tag             =   "25612"
         Top             =   45
         Width           =   1320
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00F4F0F2&
         Caption         =   "종료(&X)"
         Height          =   510
         Left            =   6000
         Style           =   1  '그래픽
         TabIndex        =   18
         Tag             =   "25612"
         Top             =   45
         Width           =   1320
      End
   End
   Begin MSComDlg.CommonDialog DlgSave 
      Left            =   5385
      Top             =   255
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblFIndx 
      BackColor       =   &H00DBE6E6&
      BackStyle       =   0  '투명
      Caption         =   "Sub Key"
      Height          =   225
      Left            =   345
      TabIndex        =   14
      Top             =   1140
      Width           =   2565
   End
   Begin VB.Label Label4 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Laboratory Information System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00613636&
      Height          =   495
      Left            =   270
      TabIndex        =   13
      Top             =   465
      Width           =   4095
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00D191A2&
      BorderWidth     =   3
      FillColor       =   &H00F7F0F0&
      FillStyle       =   0  '단색
      Height          =   495
      Left            =   270
      Shape           =   4  '둥근 사각형
      Top             =   360
      Width           =   4635
   End
   Begin VB.Label lblRName 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Laboratory Information System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00613636&
      Height          =   315
      Left            =   510
      TabIndex        =   12
      Top             =   450
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00808080&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  '단색
      Height          =   375
      Index           =   1
      Left            =   240
      Shape           =   4  '둥근 사각형
      Top             =   1035
      Width           =   2745
   End
End
Attribute VB_Name = "frm343Template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const cLableCount = 3


Private fRKey As String
Private objSql As New clsLISSqlStatement
Private WithEvents objCodeList As clsPopUpList
Attribute objCodeList.VB_VarHelpID = -1

Public Property Get Rkey() As String
    Rkey = fRKey
End Property

Public Property Let Rkey(ByVal vNewValue As String)
    fRKey = vNewValue
    cmdPopupList.Visible = False
    lbltestNm.Caption = ""
    lbltestNm.Visible = False
    Select Case Rkey
        Case LC4_TextResult, LC4_ClinicalNotice, LC4_TestItemComment
            cmdPopupList.Visible = True
            lbltestNm.Caption = ""
            lbltestNm.Visible = True
    End Select
    
    If fRKey = LC4_FootNote Then
        cmdExcel.Visible = True
    Else
        cmdExcel.Visible = False
    End If
    
End Property

Public Property Get RName() As String
    RName = lblRName
End Property

Public Property Let RName(ByVal vNewValue As String)
    lblRName = vNewValue
    
End Property

Private Sub LoadSubKey()
    
    Dim i As Integer, SSQL As String
    Dim dsSKey As Recordset

    SSQL = objSql.SqlLAB034CodeList(Rkey, "*")
    
    Set dsSKey = New Recordset
    dsSKey.Open SSQL, DBConn

    lstSubKey.Clear
    
    ClearScreen
    txtSubKey.Locked = False
    
    If dsSKey.RecordCount < 1 Then Set dsSKey = Nothing: Exit Sub

    dsSKey.MoveFirst
    For i = 1 To dsSKey.RecordCount
        Select Case Rkey
            Case LC4_TextResult, LC4_ClinicalNotice, LC4_TestItemComment
                lstSubKey.AddItem "" & dsSKey.Fields("cdval1").Value & vbTab & dsSKey.Fields("testnm").Value & ""
            Case LC4_FootNote, LC4_Remark, LC4_CancelReason, LC4_ModifyReason, _
                 LC4_QCRejReason, LC4_Calibration, LC4_AccessComment
                 lstSubKey.AddItem "" & dsSKey.Fields("cdval1").Value & vbTab & dsSKey.Fields("text1").Value & ""
            Case LC4_WarnInfect
                lstSubKey.AddItem "" & dsSKey.Fields("cdval1").Value & vbTab & dsSKey.Fields("field1").Value & ""
            Case Else
                lstSubKey.AddItem "" & dsSKey.Fields("cdval1").Value
        End Select
        dsSKey.MoveNext
    Next i

    Set dsSKey = Nothing

End Sub


Private Sub cmdClear_Click()
    lstSubKey.ListIndex = -1
    ClearScreen
    txtSubKey.Locked = False
    txtSubKey.SetFocus
End Sub

Private Sub ClearScreen()
    
    Dim i As Integer

    txtSubKey = ""
    For i = 1 To cLableCount
        txtVal(i - 1) = ""
    Next i
    lbltestNm.Caption = ""
End Sub

Private Sub cmdDelete_Click()
    
    Dim sMsg As String
    Dim sRes As Integer, sStyle As Integer

    If Trim(txtSubKey) = "" Then
        MsgBox "Index가 지정이 잘못 되었습니다. 확인 후 처리 바랍니다."
        Exit Sub
    End If

    sMsg = lblRName & " Table에서 (" & txtSubKey.Text & ") Key와 Data를 삭제합니다" & Chr$(13) & Chr$(10) & _
        "정말 삭제해도 좋습니까?"
    sStyle = vbYesNo + vbCritical + vbDefaultButton2
    
    sRes = MsgBox(sMsg, sStyle, "삭제 확인")
    If sRes = vbYes Then
        Call DeleteData
    Else
        Exit Sub
    End If

    cmdClear_Click
    
End Sub

Private Sub DeleteData()
    
    Dim SSQL As String

    SSQL = objSql.SqlDeleteLAB034(Rkey, txtSubKey.Text)
           

On Error GoTo DBExecError

    DBConn.BeginTrans
    DBConn.Execute SSQL              'dbconn.execute (sSQL)
    DBConn.CommitTrans

    'MsgBox "정상적으로 삭제 되었습니다. 다음 작업을 처리하세요", , "삭제 완료"
'    medMain.stsBar.Panels(2).Text = "정상적으로 삭제 되었습니다. 다음 작업을 처리하세요"
    LoadSubKey
    
    Exit Sub

DBExecError:
    DBConn.RollbackTrans

End Sub

Private Sub cmdExcel_Click()
    Dim strCd   As String
    Dim strNm   As String
    Dim strTmp  As String
    Dim aryNm() As String
    Dim i, j    As Integer
    
    If lstSubKey.ListCount < 1 Then Exit Sub
    
    With tblExcel
        .MaxRows = 0: .MaxCols = 2
        For i = 0 To lstSubKey.ListCount - 1
            strNm = ""
            strCd = medGetP(lstSubKey.List(i), 1, vbTab)
            strTmp = medGetP(lstSubKey.List(i), 2, vbTab)
            
            aryNm = Split(strTmp, Chr(13))
            
            For j = LBound(aryNm) To UBound(aryNm)
                If j = 0 Then
                    strNm = strNm & aryNm(j)
                Else
                    strNm = strNm & " " & Mid(aryNm(j), 2, Len(aryNm(j)))
                End If
            Next j
            
            .MaxRows = i + 1
            .Row = i + 1
            
            .Col = 1: .Value = strCd
            .Col = 2: .Value = strNm
        Next
    End With
    
    DlgSave.InitDir = "C:\"
    DlgSave.Filter = "ExCelFile(*.XLS)|*.XLS"
    DlgSave.FileName = "FootNote 내역"
    DlgSave.ShowSave
    
    tblExcel.SaveTabFile (DlgSave.FileName)
    
End Sub

Private Sub cmdExit_Click()

    'ReturnDB
'    DbClose
    Unload Me

End Sub

Private Sub cmdSave_Click()
    
    Dim i As Integer, sIndexKey As String, sFlag As String
    Dim SSQL As String, dsChk As Recordset
    Dim objsSQL As clsLISSqlCodeMaster
    
    If Trim(txtSubKey) = "" Then
        MsgBox "Index가 지정이 잘못 되었습니다. 확인 후 처리 바랍니다."
        Exit Sub
    End If

    sFlag = "I"
    
    Set objsSQL = New clsLISSqlCodeMaster
    
        
    With objsSQL
        Set dsChk = New Recordset
        dsChk.Open .GetComCdTemp(Rkey, Trim(txtSubKey.Text)), DBConn
    End With
            
    If dsChk.RecordCount = 1 Then sFlag = "U"
    
    Set dsChk = Nothing
    Set objsSQL = Nothing

    Select Case sFlag
        Case "I": Call CommonInsert
        Case "U": Call CommonUpdate
        Case Else: MsgBox "시스템에 오류가 있습니다."
    End Select
    
    cmdClear_Click
    
End Sub

Private Sub CommonInsert()
    
    Dim SSQL As String

    SSQL = objSql.SqlSaveLAB034(Rkey, Trim(txtSubKey.Text), Trim(txtVal(0).Text), _
                                Trim(txtVal(1).Text), Trim(txtVal(2).Text), 1)
    
On Error GoTo DBExecError

    DBConn.BeginTrans
    DBConn.Execute SSQL              'dbconn.execute (sSQL)
    DBConn.CommitTrans
 
    'MsgBox "정상적으로 삽입 처리 되었습니다. 다음 작업을 처리하세요", , "삽입 완료"
'    medMain.stsBar.Panels(2).Text = "정상적으로 삽입 처리 되었습니다. 다음 작업을 처리하세요"
    LoadSubKey

    Exit Sub

DBExecError:
    DBConn.RollbackTrans
 
End Sub

Private Sub CommonUpdate()
    
    Dim SSQL As String

    SSQL = objSql.SqlSaveLAB034(Rkey, Trim(txtSubKey.Text), Trim(txtVal(0).Text), _
                                Trim(txtVal(1).Text), Trim(txtVal(2).Text), 2)
           

On Error GoTo DBExecError

    DBConn.BeginTrans
    DBConn.Execute SSQL              'dbconn.execute (sSQL)
    DBConn.CommitTrans
 
    'MsgBox "정상적으로 갱신 되었습니다. 다음 작업을 처리하세요", , "갱신 완료"
'    medMain.stsBar.Panels(2).Text = "정상적으로 갱신 되었습니다. 다음 작업을 처리하세요"
    Exit Sub

DBExecError:
    DBConn.RollbackTrans

End Sub

Private Sub Form_Load()
    
    Me.WindowState = 2
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objSql = Nothing
    Set objCodeList = Nothing
End Sub

Private Sub lblRName_Change()
    LoadFieldNm
    LoadSubKey
End Sub

Private Sub LoadFieldNm()

    Dim SSQL As String
    Dim dsSKey As Recordset
    Dim strFields As String
    Dim i As Integer
    Dim objComSql As clsLISSqlCodeMaster
    Dim strCOM001_CdIndex As String
    Dim strTmp As String

    strCOM001_CdIndex = "LC4"
    
    Set objComSql = New clsLISSqlCodeMaster
    With objComSql
        Set dsSKey = New Recordset
        dsSKey.Open .GetComCdIndex(strCOM001_CdIndex, Rkey), DBConn
    End With
    
    
    SSQL = objComSql.GetComCdIndex(strCOM001_CdIndex, Rkey)
           
    Set dsSKey = New Recordset
    dsSKey.Open SSQL, DBConn
    
    If Not dsSKey.EOF Then
        strFields = "" & dsSKey.Fields("Text1").Value
    Else
        strFields = ""
    End If
    
    Set dsSKey = Nothing
    Set objComSql = Nothing
    
    For i = 0 To lblCap.Count - 1
        strTmp = medShift(strFields, ";")
        lblCap(i).Caption = medShift(strTmp, ":")
        If i > 0 Then
            If Trim(lblCap(i)) = "" Then
                txtVal(i - 1).Visible = False
            Else
                txtVal(i - 1).Visible = True
                txtVal(i - 1).MaxLength = Val(strTmp)
            End If
        End If
    Next
    
        
End Sub

Private Sub lstSubKey_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then LoadData
End Sub

Private Sub lstSubKey_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then LoadData
End Sub

Private Sub LoadData()
    
    Dim SSQL As String
    Dim dsInfo As Recordset
    
    If lstSubKey.ListIndex < 0 Then Exit Sub
    
    SSQL = objSql.SqlLAB034CodeList(Rkey, "*", medGetP(lstSubKey.List(lstSubKey.ListIndex), 1, vbTab))
           
    Set dsInfo = New Recordset
    dsInfo.Open SSQL, DBConn
    'Set dsInfo = openrecordset(sSQL, &H0&)
    
    ClearScreen
    
    If dsInfo.RecordCount < 1 Then
        MsgBox "등록되어 있지 않은 ID 입니다."
        Exit Sub
    End If
    
    Select Case Rkey
        Case LC4_TextResult, LC4_ClinicalNotice, LC4_TestItemComment
            lbltestNm.Caption = medGetP(lstSubKey.List(lstSubKey.ListIndex), 2, vbTab)
    End Select
    
    txtSubKey = "" & dsInfo.Fields("cdval1").Value
    txtVal(0) = "" & dsInfo.Fields("field1").Value
    txtVal(1) = "" & dsInfo.Fields("text1").Value
    txtVal(2) = "" & dsInfo.Fields("text2").Value

    ' 데이타 읽고 나서 키를 바꿀수 없게..
    txtSubKey.Locked = chkKeyLock.Value
'    medMain.stsBar.Panels(2).Text = "정상적으로 조회 되었습니다."

    Set dsInfo = Nothing

End Sub

Private Sub txtSubKey_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtSubKey_LostFocus()
    Dim SSQL As String
    Dim RS   As Recordset
    
    lbltestNm.Caption = ""
    
    Select Case Rkey
        Case LC4_TextResult, LC4_ClinicalNotice, LC4_TestItemComment
            SSQL = GetTestNm(Trim(txtSubKey.Text))
    End Select
    
    If SSQL = "" Then Exit Sub
    Set RS = New Recordset
    RS.Open SSQL, DBConn
    
    If Not RS.EOF Then
        lbltestNm.Caption = RS.Fields("testnm").Value & ""
    End If
    
    Set RS = Nothing
    
End Sub

Private Function GetTestNm(Optional ByVal sTestCd As String = "") As String
    Dim SSQL As String

    Select Case Rkey
        Case LC4_TextResult
            SSQL = " select a.testcd,a.testnm from " & T_LAB001 & " a" & _
                   " where a.txttype in('1','2')" & _
                   " and  a.applydt = ( select max(applydt) from " & T_LAB001 & _
                  "                     where testcd = a.testcd )"
            If sTestCd <> "" Then SSQL = SSQL & " and " & DBW("a.testcd=", sTestCd)
        Case LC4_ClinicalNotice, LC4_TestItemComment
            SSQL = " select a.testcd,a.testnm from " & T_LAB001 & " a" & _
                   " where  a.applydt = ( select max(applydt) from " & T_LAB001 & _
                   "                     where testcd = a.testcd )"
            If sTestCd <> "" Then SSQL = SSQL & " and " & DBW("a.testcd=", sTestCd)
    End Select
    GetTestNm = SSQL
End Function
'% List 버튼을 클릭한 경우 코드리스트를 팝업한다.
Private Sub cmdPopupList_Click()

    Dim tmpSql  As String
    Dim lngTop  As Long
    Dim lngLeft As Long

    Set objCodeList = New clsPopUpList
    With objCodeList
        lngTop = txtSubKey.Top + 2350
        lngLeft = Me.Left + Frame1.Left + txtSubKey.Left + 50
        .Connection = DBConn
        .Tag = "TestCd"
        .FormCaption = "검사항목 리스트"
        .ColumnHeaderText = "검사코드;검사명"
        tmpSql = GetTestNm
        .LoadPopUp tmpSql ', lngTop, lngLeft
        txtSubKey.Text = medGetP(.SelectedString, 1, ";")
        lbltestNm.Caption = medGetP(.SelectedString, 2, ";")
    End With
    Set objCodeList = Nothing
End Sub
