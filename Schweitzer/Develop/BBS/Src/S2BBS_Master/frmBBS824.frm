VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmBBS824 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   3  'Å©±â °íÁ¤ ´ëÈ­ »óÀÚ
   Caption         =   "ÇåÇ÷½ÇÀûÅëÁö¼­"
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
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00F4F0F2&
      Caption         =   "ÀúÀå(&S)"
      Height          =   420
      Left            =   4800
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   6
      Top             =   7740
      Width           =   1260
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "Á¾·á(&X)"
      Height          =   420
      Left            =   6120
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   5
      Top             =   7740
      Width           =   1260
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "È­¸éÁö¿ò(&C)"
      Height          =   420
      Left            =   3480
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   4
      Top             =   7740
      Width           =   1260
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   375
      Left            =   1140
      TabIndex        =   2
      Top             =   4500
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   661
      BackColor       =   8421504
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   " Á¦Á¦ ¼±ÅÃ"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   375
      Left            =   1140
      TabIndex        =   1
      Top             =   480
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   661
      BackColor       =   8421504
      ForeColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   " Ãâ·Â ¾ç½Ä"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblResult 
      Height          =   3405
      Left            =   1140
      TabIndex        =   0
      Top             =   900
      Width           =   8430
      _Version        =   196608
      _ExtentX        =   14870
      _ExtentY        =   6006
      _StockProps     =   64
      AutoSize        =   -1  'True
      BackColorStyle  =   1
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "µ¸¿ò"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14411494
      GridShowVert    =   0   'False
      MaxCols         =   8
      MaxRows         =   8
      OperationMode   =   1
      ScrollBars      =   0
      ShadowColor     =   14737632
      ShadowDark      =   14737632
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS824.frx":0000
      TextTip         =   4
   End
   Begin FPSpread.vaSpread tblCompo 
      Height          =   2430
      Left            =   1140
      TabIndex        =   3
      Tag             =   "10114"
      Top             =   4920
      Width           =   8430
      _Version        =   196608
      _ExtentX        =   14870
      _ExtentY        =   4286
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      ButtonDrawMode  =   4
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormulaSync     =   0   'False
      GridShowVert    =   0   'False
      MaxCols         =   4
      MaxRows         =   8
      MoveActiveOnFocus=   0   'False
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      ShadowColor     =   14737632
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "frmBBS824.frx":138D
      StartingColNumber=   2
      VirtualRows     =   24
      VisibleRows     =   10
   End
End
Attribute VB_Name = "frmBBS824"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdClear_Click()
    Call QueryComponent
    Call QueryPrintForm
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Save = True Then
        Call QueryComponent
        Call QueryPrintForm
    End If
End Sub

Private Function Save() As Boolean
    Dim cdval1(2) As String
    Dim field1(2) As String
    Dim field2(2) As String
    Dim objcom003 As clsCom003
    
    Dim Row As Long
    Dim CompoCd As String
    Dim idx As Long
    
    cdval1(0) = "01": field1(0) = "Ç÷  Àå": field2(0) = ""
    cdval1(1) = "02": field1(1) = "Ç÷¼ÒÆÇ": field2(1) = ""
    cdval1(2) = "03": field1(2) = "¹éÇ÷±¸": field2(2) = ""
    
    With tblCompo
        For Row = 1 To .MaxRows
            .Row = Row
            .Col = 1
            CompoCd = .Value
            
            If CompoCd = "" Then Exit For
            
            .Col = 4
            idx = .Value
            
            If field2(idx) = "" Then
                field2(idx) = CompoCd
            Else
                field2(idx) = field2(idx) & "," & CompoCd
            End If
        Next Row
    End With
    
    Set objcom003 = New clsCom003
    
On Error GoTo Save_error

    DBConn.BeginTrans
    
    For idx = 0 To 2
    
        objcom003.CDINDEX = BC2_PHERESIS_COMPO
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

Private Sub Form_Load()
    Call QueryComponent
    Call QueryPrintForm
End Sub

Private Sub QueryComponent()
    Dim objCompo As clsComponent
    Dim DrRS As Recordset
    Dim i As Long
    
    Dim Row As Long
    Dim CompoCd As String
    Dim componm As String
    Dim abbrnm As String
    
    Set objCompo = New clsComponent
    Set DrRS = objCompo.GetList
    Set objCompo = Nothing
    
    
    If DrRS Is Nothing Then
        MsgBox "µî·ÏµÈ Ç÷¾×Á¦Á¦°¡ ¾ø½À´Ï´Ù.", vbCritical, Me.Caption
        Exit Sub
    End If
    
    medClearTable tblCompo
    
    With DrRS
        Row = 0
        For i = 1 To .RecordCount
            If .Fields("pherefg").Value & "" = "1" Then
                CompoCd = .Fields("compocd").Value & ""
                componm = .Fields("componm").Value & ""
                abbrnm = .Fields("abbrnm").Value & ""
                
                Row = Row + 1
                
                With tblCompo
                    If Row > .MaxRows Then .MaxRows = .MaxRows + 1
                    .Row = Row
                    .Col = 1: .Value = CompoCd
                    .Col = 2: .Value = componm
                    .Col = 3: .Value = abbrnm
                    .Col = 4: .Value = -1
                End With
            End If
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
    Dim CompoCd As String
    Dim Row As Long
    
    Dim cdval1 As String
    Dim field1 As String
    Dim field2 As String
    
    Set objcom003 = New clsCom003
    Set DrRS = objcom003.OpenRecordSet(BC2_PHERESIS_COMPO)
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
                CompoCd = medGetP(field2, j, ",")
                If CompoCd = "" Then Exit Do
                
                With tblCompo
                    For Row = 1 To .MaxRows
                        .Row = Row
                        .Col = 1
                        If .Value = CompoCd Then
                            If cdval1 = "01" Then
                                .Col = 4: .Value = 0
                            ElseIf cdval1 = "02" Then
                                .Col = 4: .Value = 1
                            ElseIf cdval1 = "03" Then
                                .Col = 4: .Value = 2
                            End If
                        End If
                    Next Row
                End With
            Loop
            
            .MoveNext
        Next i
    End With
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
        
        If Row = 3 Then
            If Col = 6 Or Col = 7 Or Col = 8 Then
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
