VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmBBS825 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "부적격혈액이송"
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
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "화면지움(&C)"
      Height          =   420
      Left            =   3360
      Style           =   1  '그래픽
      TabIndex        =   6
      Top             =   7500
      Width           =   1260
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종료(&X)"
      Height          =   420
      Left            =   6000
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   7500
      Width           =   1260
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "저장(&S)"
      Height          =   420
      Left            =   4680
      Style           =   1  '그래픽
      TabIndex        =   4
      Top             =   7500
      Width           =   1260
   End
   Begin FPSpread.vaSpread tblResult 
      Height          =   1395
      Left            =   1860
      TabIndex        =   0
      Top             =   1020
      Width           =   6975
      _Version        =   196608
      _ExtentX        =   12312
      _ExtentY        =   2461
      _StockProps     =   64
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14411494
      MaxCols         =   11
      MaxRows         =   5
      OperationMode   =   1
      ScrollBars      =   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS825.frx":0000
      TextTip         =   4
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   375
      Left            =   1860
      TabIndex        =   1
      Top             =   600
      Width           =   6975
      _ExtentX        =   12303
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
      Caption         =   " 출력 양식"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblReason 
      Height          =   4110
      Left            =   1860
      TabIndex        =   2
      Tag             =   "10114"
      Top             =   3060
      Width           =   6975
      _Version        =   196608
      _ExtentX        =   12303
      _ExtentY        =   7250
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
      MaxCols         =   3
      MaxRows         =   15
      MoveActiveOnFocus=   0   'False
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS825.frx":0D71
      StartingColNumber=   2
      VirtualRows     =   24
      VisibleRows     =   8
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   375
      Left            =   1860
      TabIndex        =   3
      Top             =   2640
      Width           =   6975
      _ExtentX        =   12303
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
      Caption         =   " 사유 선택"
      Appearance      =   0
   End
End
Attribute VB_Name = "frmBBS825"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    Call QueryReason
    Call QueryPrintForm
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Save = True Then
        Call QueryReason
        Call QueryPrintForm
    End If
End Sub

Private Sub Form_Load()
    Call QueryReason
    Call QueryPrintForm
End Sub

Private Sub tblResult_Click(ByVal Col As Long, ByVal Row As Long)
    Static ClickRow As Long
    Static ClickCol As Long
    
    With tblResult
        .Row = Row
        .Col = Col
        .Action = ActionActiveCell
        
        .Row = ClickRow
        .Col = ClickCol
        .BackColor = RGB(255, 255, 255)
        .ForeColor = RGB(0, 0, 0)
        
        If Row = 2 Then
            If Col >= 3 And Col <= 8 Then
                .Row = Row
                .Col = Col
                .BackColor = RGB(0, 0, 150)
                .ForeColor = RGB(255, 255, 255)
            End If
        End If
        
        ClickRow = Row
        ClickCol = Col
    End With

End Sub



Private Sub QueryReason()
    Dim objcom003 As clsCom003
    Dim DrRS As Recordset
    Dim i As Long
    
    Dim Row As Long
    Dim code As String
    Dim name As String
    
    Set objcom003 = New clsCom003
    Set DrRS = objcom003.OpenRecordSet(BC2_DONOR_REASON)
    Set objcom003 = Nothing
    
    
    If DrRS Is Nothing Then
        MsgBox "등록된 부적격사유가 없습니다.", vbCritical, Me.Caption
        Exit Sub
    End If
    
    medClearTable tblReason
    
    With DrRS
        Row = 0
        For i = 1 To .RecordCount
            code = .Fields("cdval1").Value & ""
            name = .Fields("field1").Value & ""
            
            Row = Row + 1
            
            With tblReason
                If Row > .MaxRows Then .MaxRows = .MaxRows + 1
                .Row = Row
                .Col = 1: .Value = code
                .Col = 2: .Value = name
                .Col = 3: .Value = -1
            End With
            .MoveNext
        Next i
    End With
    
    Set DrRS = Nothing
End Sub

Private Sub QueryPrintForm()
    Dim objcom003 As clsCom003
    Dim DrRS As Recordset
    Dim i As Long
    Dim j As Long
    Dim code As String
    Dim Row As Long
    
    Dim cdval1 As String
    Dim field1 As String
    Dim field2 As String
    
    Set objcom003 = New clsCom003
    Set DrRS = objcom003.OpenRecordSet(BC2_UNACCEPT_REASON)
    Set objcom003 = Nothing

    If DrRS Is Nothing Then Exit Sub

    With DrRS
        For i = 1 To .RecordCount
            cdval1 = .Fields("cdval1").Value & ""
            field1 = .Fields("field1").Value & ""
            field2 = .Fields("field2").Value & ""
            
            j = 0
            Do
                j = j + 1
                code = medGetP(field2, j, ",")
                If code = "" Then Exit Do
                
                With tblReason
                    For Row = 1 To .MaxRows
                        .Row = Row
                        .Col = 1
                        If .Value = code Then
                            Select Case cdval1
                                Case "01": .Col = 3: .Value = 0
                                Case "02": .Col = 3: .Value = 1
                                Case "03": .Col = 3: .Value = 2
                                Case "04": .Col = 3: .Value = 3
                                Case "05": .Col = 3: .Value = 4
                                Case "06": .Col = 3: .Value = 5
                            End Select
                        End If
                    Next Row
                End With
            Loop
            
            .MoveNext
        Next i
    End With
End Sub

Private Function Save() As Boolean
    Dim cdval1(5) As String
    Dim field1(5) As String
    Dim field2(5) As String
    Dim objcom003 As clsCom003
    
    Dim Row As Long
    Dim code As String
    Dim idx As Long
    
    cdval1(0) = "01": field1(0) = "ALT":     field2(0) = ""
    cdval1(1) = "02": field1(1) = "B형간염": field2(1) = ""
    cdval1(2) = "03": field1(2) = "C형간염": field2(2) = ""
    cdval1(3) = "04": field1(3) = "매독":    field2(3) = ""
    cdval1(4) = "05": field1(4) = "에이즈":  field2(4) = ""
    cdval1(5) = "06": field1(5) = "기타":    field2(5) = ""
    
    With tblReason
        For Row = 1 To .MaxRows
            .Row = Row
            .Col = 1
            code = .Value
            
            If code = "" Then Exit For
            
            .Col = 3
            idx = .Value
            
            If field2(idx) = "" Then
                field2(idx) = code
            Else
                field2(idx) = field2(idx) & "," & code
            End If
        Next Row
    End With
    
    Set objcom003 = New clsCom003
    
On Error GoTo Save_error

    DBConn.BeginTrans
    
    For idx = 0 To 5
    
        objcom003.CDINDEX = BC2_UNACCEPT_REASON
        objcom003.cdval1 = cdval1(idx)
        objcom003.field1 = field1(idx)
        objcom003.field2 = field2(idx)
        
        If objcom003.Save() = False Then GoTo Save_error
    Next idx
    
    DBConn.CommitTrans
    Set objcom003 = Nothing
    Save = True
    Exit Function
    
Save_error:

    DBConn.RollbackTrans
    Set objcom003 = Nothing
    Save = False
End Function
