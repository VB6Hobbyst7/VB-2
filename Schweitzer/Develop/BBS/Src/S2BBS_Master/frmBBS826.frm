VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBBS826 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "수가발생코드등록"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   420
      Left            =   8160
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   7860
      Width           =   1260
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   420
      Left            =   6840
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   7860
      Width           =   1260
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   661
      BackColor       =   8421504
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   " 수혈 코드"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblTestCd 
      Height          =   7710
      Left            =   240
      TabIndex        =   1
      Tag             =   "10114"
      Top             =   600
      Width           =   5115
      _Version        =   196608
      _ExtentX        =   9022
      _ExtentY        =   13600
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      ButtonDrawMode  =   4
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormulaSync     =   0   'False
      GridShowVert    =   0   'False
      MaxCols         =   2
      MaxRows         =   30
      MoveActiveOnFocus=   0   'False
      OperationMode   =   3
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS826.frx":0000
      StartingColNumber=   2
      VirtualRows     =   24
      VisibleRows     =   8
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   180
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   661
      BackColor       =   8421504
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   " 수가 발생 코드"
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   7215
      Left            =   5520
      TabIndex        =   3
      Top             =   510
      Width           =   5115
      Begin VB.CheckBox chkVolumeDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "용량에 관계없이 동일한 코드 적용"
         Enabled         =   0   'False
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   1740
         Width           =   3075
      End
      Begin VB.CheckBox chkNewDiv 
         BackColor       =   &H00DBE6E6&
         Caption         =   "코드발생 안시킴"
         Enabled         =   0   'False
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   1380
         Width           =   1635
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00F4F0F2&
         Caption         =   "저장(&S)"
         Enabled         =   0   'False
         Height          =   420
         Left            =   3660
         Style           =   1  '그래픽
         TabIndex        =   8
         Top             =   1200
         Width           =   1260
      End
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H00F4F0F2&
         Caption         =   "신규등록"
         Height          =   420
         Left            =   3660
         Style           =   1  '그래픽
         TabIndex        =   7
         Top             =   720
         Width           =   1260
      End
      Begin MSComctlLib.TabStrip tabApplyDt 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "2001-01-01"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin FPSpread.vaSpread tblNewTest 
         Height          =   1950
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Tag             =   "10114"
         Top             =   2580
         Width           =   4575
         _Version        =   196608
         _ExtentX        =   8070
         _ExtentY        =   3440
         _StockProps     =   64
         Enabled         =   0   'False
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         ButtonDrawMode  =   4
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         EditModePermanent=   -1  'True
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GridShowVert    =   0   'False
         MaxCols         =   3
         MaxRows         =   50
         MoveActiveOnFocus=   0   'False
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS826.frx":051E
         StartingColNumber=   2
         VirtualRows     =   24
         VisibleRows     =   8
      End
      Begin FPSpread.vaSpread tblNewTest 
         Height          =   1950
         Index           =   1
         Left            =   300
         TabIndex        =   14
         Tag             =   "10114"
         Top             =   5040
         Width           =   4575
         _Version        =   196608
         _ExtentX        =   8070
         _ExtentY        =   3440
         _StockProps     =   64
         Enabled         =   0   'False
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   1
         ButtonDrawMode  =   4
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         EditModePermanent=   -1  'True
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GridShowVert    =   0   'False
         MaxCols         =   3
         MaxRows         =   50
         MoveActiveOnFocus=   0   'False
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "frmBBS826.frx":0BA0
         StartingColNumber=   2
         VirtualRows     =   24
         VisibleRows     =   8
      End
      Begin MSComCtl2.DTPicker dtpApplyDt 
         Height          =   315
         Left            =   1320
         TabIndex        =   15
         Top             =   840
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   62849027
         CurrentDate     =   36956
      End
      Begin VB.Label lblApplyDt 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "적용일자"
         Height          =   180
         Left            =   480
         TabIndex        =   16
         Top             =   900
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblVolume 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "400mL"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   540
         TabIndex        =   12
         Top             =   4800
         Width           =   630
      End
      Begin VB.Label lblVolume 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "320mL"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   480
         TabIndex        =   11
         Top             =   2340
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmBBS826"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Enum TblColumn
    tcNEWTESTCD = 1
    tcNEWTESTNM
    tcONCEDIV
End Enum



Private Sub chkNewDiv_Click()
    If chkNewDiv.Value = 1 Then
        tblNewTest(0).Enabled = False
        tblNewTest(1).Enabled = False
    Else
        tblNewTest(0).Enabled = True
        tblNewTest(1).Enabled = True
    End If
End Sub

Private Sub chkVolumeDiv_Click()
    If chkVolumeDiv.Value = 1 Then
        lblVolume(0).Caption = "용량에 관계없이"
        tblNewTest(0).Enabled = True
        lblVolume(1).Caption = "입력안함"
        tblNewTest(1).Enabled = False
    Else
        lblVolume(0).Caption = "320mL"
        tblNewTest(0).Enabled = True
        lblVolume(1).Caption = "400mL"
        tblNewTest(1).Enabled = True
    End If
End Sub

Private Sub cmdClear_Click()
    Call Clear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()

    If tabApplyDt.Tabs.Count < 1 Then
        dtpApplyDt.MinDate = GetSystemDate
    Else
        dtpApplyDt.MinDate = CDate(tabApplyDt.Tabs.Item(1).Caption) + 1
    End If


    '신규입력준비
    dtpApplyDt = GetSystemDate
    lblApplyDt.Visible = True
    dtpApplyDt.Visible = True
    
    chkNewDiv.Enabled = True
    chkVolumeDiv.Enabled = True
    tblNewTest(0).Enabled = True
    tblNewTest(1).Enabled = True
    cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()
    If Save = True Then
    End If
End Sub

Private Sub Form_Load()
    Call QueryTestCds
End Sub

Private Function GetBBS001(ByVal dt As String) As Recordset
    '수혈처방을 불러온다.
    Dim strSql As String
    
    strSql = " SELECT testcd,testnm,testdiv " & _
             " FROM  " & T_BBS001 & " a" & _
             " WHERE a.applydt = ( SELECT max(b.applydt) " & _
                                  " FROM " & T_BBS001 & " b " & _
                                  " WHERE a.testcd = b.testcd " & _
                                  " AND   " & DBW("b.applydt<=", dt) & ") " & _
             " ORDER BY testcd "
    
    Set GetBBS001 = New Recordset
    GetBBS001.Open strSql, DBConn
    
End Function


Private Sub QueryTestCds()
    Dim RS          As Recordset
    Dim code        As String
    Dim name        As String
    Dim i           As Long
    
    Call medClearTable(tblTestCd)

    Set RS = GetBBS001(Format(GetSystemDate, PRESENTDATE_FORMAT))
    
'    If RS.DBerror Then
'        'dbconn.DisplayErrors
'        Set RS = Nothing
'        Exit Sub
'    End If

    With RS
        For i = 1 To .RecordCount
            code = .Fields("testcd").Value & ""
            name = .Fields("testnm").Value & ""
            
            With tblTestCd
                .Row = i
                If .Row > .MaxRows Then .MaxRows = .MaxRows + 1
                
                .Col = 1: .Value = code
                .Col = 2: .Value = name
            End With
            
            .MoveNext
        Next i
    End With

    Set RS = Nothing
End Sub

Private Sub tabApplyDt_Click()
    Dim TestCd As String
    Dim ApplyDt As String
    Dim Volume As Long
    Dim newcode As String
    Dim newname As String
    Dim oncediv As String
    
    Dim objNewTestCode As clsNewTestCode
    Dim DrRS As Recordset
    Dim i As Long
    
    With tblTestCd
        .Row = .ActiveRow
        .Col = 1
        TestCd = .Value
    End With
    ApplyDt = Format(tabApplyDt.SelectedItem.Caption, PRESENTDATE_FORMAT)
    
    Set objNewTestCode = New clsNewTestCode
    
    '헤더 읽기
    Set DrRS = objNewTestCode.GetHeader(TestCd, ApplyDt)
    If DrRS Is Nothing Then Exit Sub
    
    With DrRS
        If .RecordCount > 0 Then
            chkNewDiv.Value = .Fields("newdiv").Value & ""
            chkVolumeDiv.Value = .Fields("volumediv").Value & ""
        End If
    End With
    Set DrRS = Nothing
    
    '바디 읽기
    If chkVolumeDiv.Value = 1 Then
        '용량에 관계없이....
        Set DrRS = objNewTestCode.GetBody(TestCd, ApplyDt, -1)
        If DrRS Is Nothing Then Exit Sub
        
        With DrRS
            If .RecordCount > 0 Then
                For i = 1 To .RecordCount
                    newcode = .Fields("newtestcd").Value & ""
                    newname = .Fields("newtestnm").Value & ""
                    oncediv = .Fields("oncediv").Value & ""
                    
                    With tblNewTest(0)
                        .Row = i
                        .Col = 1: .Value = newcode
                        .Col = 2: .Value = newname
                        .Col = 3: .Value = Val(oncediv)
                    End With
                    
                    .MoveNext
                Next i
            End If
        End With
        Set DrRS = Nothing
    Else
        '320mL
        Set DrRS = objNewTestCode.GetBody(TestCd, ApplyDt, 320)
        If DrRS Is Nothing Then Exit Sub
        
        With DrRS
            If .RecordCount > 0 Then
                For i = 1 To .RecordCount
                    newcode = .Fields("newtestcd").Value & ""
                    newname = .Fields("newtestnm").Value & ""
                    oncediv = .Fields("oncediv").Value & ""
                    
                    With tblNewTest(0)
                        .Row = i
                        .Col = 1: .Value = newcode
                        .Col = 2: .Value = newname
                        .Col = 3: .Value = Val(oncediv)
                    End With
                    
                    .MoveNext
                Next i
            End If
        End With
        Set DrRS = Nothing
        
        '400mL
        Set DrRS = objNewTestCode.GetBody(TestCd, ApplyDt, 400)
        If DrRS Is Nothing Then Exit Sub
        
        With DrRS
            If .RecordCount > 0 Then
                For i = 1 To .RecordCount
                    newcode = .Fields("newtestcd").Value & ""
                    newname = .Fields("newtestnm").Value & ""
                    oncediv = .Fields("oncediv").Value & ""
                    
                    With tblNewTest(1)
                        .Row = i
                        .Col = 1: .Value = newcode
                        .Col = 2: .Value = newname
                        .Col = 3: .Value = Val(oncediv)
                    End With
                    
                    .MoveNext
                Next i
            End If
        End With
        Set DrRS = Nothing
    End If
End Sub

Private Sub Clear()
    tabApplyDt.Tabs.Clear
    lblApplyDt.Visible = False
    dtpApplyDt.Visible = False
    
    chkNewDiv.Value = 0
    chkVolumeDiv.Value = 0
    medClearTable tblNewTest(0)
    tblNewTest(0).Enabled = False
    medClearTable tblNewTest(1)
    tblNewTest(1).Enabled = False
    
    
    cmdNew.Enabled = True
    cmdSave.Enabled = False
End Sub

Private Sub tblTestCd_Click(ByVal Col As Long, ByVal Row As Long)
    Dim objNewTestCode As clsNewTestCode
    Dim DrRS As Recordset
    Dim TestCd As String
    Dim aApplyDt() As Date
    Dim i As Long
    
    Clear
    
    With tblTestCd
        .Row = Row
        .Col = 1
        TestCd = .Value
    End With
    
    Set objNewTestCode = New clsNewTestCode
    If objNewTestCode.GetApplyDtList(TestCd, aApplyDt) = False Then
        Set objNewTestCode = Nothing
        Exit Sub
    End If
    
    For i = LBound(aApplyDt) To UBound(aApplyDt)
        Call tabApplyDt.Tabs.Add(, , Format(aApplyDt(i), "YYYY-MM-DD"))
    Next i
    
    Set objNewTestCode = Nothing
    
    tabApplyDt.Tabs.Item(1).Selected = True
End Sub

Private Function Save() As Boolean
    Dim TestCd As String
    Dim ApplyDt As String
    Dim newdiv As String
    Dim volumediv As String
    
    Dim Volume As Long
    Dim newcode As String
    Dim newname As String
    Dim oncediv As String
    Dim retdiv As String
    
    Dim objNewTestCode As clsNewTestCode
    Dim i As Long
    
    Set objNewTestCode = New clsNewTestCode
    
    
    With tblTestCd
        .Row = .ActiveRow
        .Col = 1
        TestCd = .Value
    End With
    ApplyDt = Format(dtpApplyDt, PRESENTDATE_FORMAT)
    newdiv = chkNewDiv.Value
    volumediv = chkVolumeDiv.Value

On Error GoTo Save_error

    DBConn.BeginTrans

    If objNewTestCode.SaveHeader(TestCd, ApplyDt, newdiv, volumediv) = False Then GoTo Save_error
    
    If newdiv = "0" Then
        If volumediv = "1" Then
            '용량관계없음
            Volume = -1
            With tblNewTest(0)
                For i = 1 To .DataRowCnt
                    .Row = i
                    .Col = 1: newcode = .Value
                    .Col = 2: newname = .Value
                    .Col = 3: oncediv = .Value
                              retdiv = IIf(oncediv = "1", "0", "1")
                              
                    If newcode = "" Then Exit For
                    
                    If objNewTestCode.SaveBody(TestCd, ApplyDt, Volume, newcode, newname, oncediv, retdiv) = False Then GoTo Save_error
                Next i
            End With
        Else
            '320mL
            Volume = 320
            With tblNewTest(0)
                For i = 1 To .DataRowCnt
                    .Row = i
                    .Col = 1: newcode = .Value
                    .Col = 2: newname = .Value
                    .Col = 3: oncediv = .Value
                              retdiv = IIf(oncediv = "1", "0", "1")
                              
                    If newcode = "" Then Exit For
                    
                    If objNewTestCode.SaveBody(TestCd, ApplyDt, Volume, newcode, newname, oncediv, retdiv) = False Then GoTo Save_error
                Next i
            End With
            '400mL
            Volume = 400
            With tblNewTest(1)
                For i = 1 To .DataRowCnt
                    .Row = i
                    .Col = 1: newcode = .Value
                    .Col = 2: newname = .Value
                    .Col = 3: oncediv = .Value
                              retdiv = IIf(oncediv = "1", "0", "1")
                              
                    If newcode = "" Then Exit For
                    
                    If objNewTestCode.SaveBody(TestCd, ApplyDt, Volume, newcode, newname, oncediv, retdiv) = False Then GoTo Save_error
                Next i
            End With
        End If
    End If
    
    
    DBConn.CommitTrans
    Save = True
    Exit Function
    
Save_error:

    DBConn.RollbackTrans
    Save = False
End Function
