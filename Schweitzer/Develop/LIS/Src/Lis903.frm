VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.OCX"
Begin VB.Form frm253MReading 
   BackColor       =   &H00DBE6E6&
   Caption         =   "배양 양성자 출력"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14670
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   14670
   Tag             =   "25200"
   WindowState     =   2  '최대화
   Begin VB.CommandButton CmdPrint 
      BackColor       =   &H00F4F0F2&
      Caption         =   "출  력"
      Height          =   480
      Left            =   11580
      Style           =   1  '그래픽
      TabIndex        =   22
      Tag             =   "25206"
      Top             =   8580
      Width           =   1425
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00F4F0F2&
      Caption         =   "&Clear"
      Height          =   480
      Left            =   10080
      Style           =   1  '그래픽
      TabIndex        =   21
      Tag             =   "25206"
      Top             =   8580
      Width           =   1425
   End
   Begin VB.CommandButton cmdWSList 
      BackColor       =   &H00DEDBDD&
      Caption         =   "▼"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9855
      Style           =   1  '그래픽
      TabIndex        =   9
      Top             =   150
      Width           =   345
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00F4F0F2&
      Caption         =   "종 료(&X)"
      Height          =   480
      Left            =   13080
      Style           =   1  '그래픽
      TabIndex        =   4
      Tag             =   "128"
      Top             =   8565
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   7230
      Left            =   180
      TabIndex        =   3
      Top             =   1275
      Width           =   14340
      Begin VB.CommandButton cmdEx1 
         BackColor       =   &H00CDE7FA&
         Caption         =   ">>"
         Height          =   350
         Left            =   9465
         Style           =   1  '그래픽
         TabIndex        =   7
         Top             =   2670
         Width           =   550
      End
      Begin VB.CommandButton cmdIn1 
         BackColor       =   &H00CDE7FA&
         Caption         =   "<<"
         Height          =   350
         Left            =   9465
         Style           =   1  '그래픽
         TabIndex        =   6
         Top             =   3150
         Width           =   550
      End
      Begin FPSpread.vaSpread ssTable 
         Height          =   6525
         Left            =   180
         TabIndex        =   5
         Tag             =   "25211"
         Top             =   600
         Width           =   9240
         _Version        =   196608
         _ExtentX        =   16298
         _ExtentY        =   11509
         _StockProps     =   64
         AutoCalc        =   0   'False
         BackColorStyle  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GridShowVert    =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   8
         MoveActiveOnFocus=   0   'False
         OperationMode   =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "Lis903.frx":0000
         UserResize      =   0
         VisibleCols     =   6
         VisibleRows     =   500
      End
      Begin FPSpread.vaSpread ssHTable 
         Height          =   6495
         Left            =   10050
         TabIndex        =   23
         Tag             =   "25211"
         Top             =   600
         Width           =   4095
         _Version        =   196608
         _ExtentX        =   7223
         _ExtentY        =   11456
         _StockProps     =   64
         AutoCalc        =   0   'False
         DisplayColHeaders=   0   'False
         EditModePermanent=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GridShowVert    =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   10
         MoveActiveOnFocus=   0   'False
         OperationMode   =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "Lis903.frx":1E7B
         UserResize      =   0
         VisibleCols     =   6
         VisibleRows     =   500
      End
      Begin VB.Label lblHCount 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00DBE6E6&
         Caption         =   "000"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   13680
         TabIndex        =   14
         Top             =   315
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblCount 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00DBE6E6&
         Caption         =   "000"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   8190
         TabIndex        =   13
         Top             =   315
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label11 
         BackColor       =   &H00DBE6E6&
         Caption         =   "배양 양성자 리스트"
         Height          =   225
         Left            =   10065
         TabIndex        =   12
         Top             =   300
         Width           =   2235
      End
      Begin VB.Label Label9 
         BackColor       =   &H00DBE6E6&
         Caption         =   "배양 대상 리스트"
         Height          =   225
         Left            =   300
         TabIndex        =   11
         Top             =   315
         Width           =   2505
      End
      Begin VB.Line Line2 
         BorderStyle     =   3  '점
         X1              =   9735
         X2              =   9735
         Y1              =   525
         Y2              =   6960
      End
   End
   Begin VB.TextBox txtWSUnit 
      Alignment       =   2  '가운데 맞춤
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8340
      TabIndex        =   1
      Text            =   "19990005"
      Top             =   150
      Width           =   1485
   End
   Begin VB.ComboBox cboWSCode 
      BackColor       =   &H00F1F5F4&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Lis903.frx":3D55
      Left            =   6885
      List            =   "Lis903.frx":3D57
      Style           =   2  '드롭다운 목록
      TabIndex        =   0
      Top             =   150
      Width           =   1470
   End
   Begin VB.ListBox lstWSUnit 
      BackColor       =   &H00F7FFF7&
      Height          =   2220
      Left            =   8340
      TabIndex        =   10
      Top             =   525
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '투명
      Caption         =   "총 검체수"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   585
      TabIndex        =   20
      Tag             =   "25203"
      Top             =   885
      Width           =   885
   End
   Begin VB.Label lblTCount 
      BackStyle       =   0  '투명
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   1575
      TabIndex        =   19
      Top             =   855
      Width           =   555
   End
   Begin VB.Label lblRcvDT 
      BackStyle       =   0  '투명
      Caption         =   "Feb 03 1999 10:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   4005
      TabIndex        =   18
      Top             =   855
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "접수 마감일/시"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2535
      TabIndex        =   17
      Tag             =   "25203"
      Top             =   900
      Width           =   1320
   End
   Begin VB.Label lblTBuiltDate 
      BackStyle       =   0  '투명
      Caption         =   "Worksheet 작성일/시"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   16
      Tag             =   "25202"
      Top             =   900
      Width           =   1965
   End
   Begin VB.Label lblBltDate 
      BackStyle       =   0  '투명
      Caption         =   "Feb 03 1999 10:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   8310
      TabIndex        =   15
      Top             =   855
      Width           =   1965
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00F1F5F5&
      BackStyle       =   1  '투명하지 않음
      Height          =   525
      Left            =   165
      Top             =   720
      Width           =   14325
   End
   Begin VB.Label lblInsResult 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "Insert Result"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   405
      TabIndex        =   8
      Tag             =   "25205"
      Top             =   7965
      Width           =   1620
   End
   Begin VB.Label lblWSUnit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Work Sheet Unit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4785
      TabIndex        =   2
      Tag             =   "25201"
      Top             =   195
      Width           =   1845
   End
End
Attribute VB_Name = "frm253MReading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objRstDic As New clsDictionary
Private objMicRst As New clsLISMicResult
Private objMicCul As New clsLISMicCulture

Private fWorkSheet() As tpMicWorkSheet
Private fNGCode() As String

Private Const fSCItem = &H8080FF          ' Worksheet List 에서 선택된 Lab-No
Private fGCItem As Long

Private Sub cboWSCode_Click()
    
    Dim i As Integer
    
    If cboWSCode.ListIndex < 0 Then Exit Sub
    
    txtWSUnit = ""
    lstWSUnit.Clear
    lstWSUnit.Visible = False
    txtWSUnit.SetFocus

    lblTCount = "": lblRcvDT = "": lblBltDate = ""
    lblCount = "": ssTable.MaxRows = 0
    
End Sub



Private Sub cmdClear_Click()
    cboWSCode.ListIndex = -1:
    txtWSUnit = ""
    ssHTable.MaxRows = 0
    Call ScreenClear
End Sub

Private Sub cmdPrint_Click()
    
    Dim MyReport As clsWorkListM
    Dim pParas As String
    Dim i As Integer
    Dim svWsCd As String, svWsUnit As String
    
    If ssHTable.MaxRows <= 0 Then Exit Sub
    
    pParas = "": svWsCd = "": svWsUnit = ""
    For i = 1 To ssHTable.MaxRows
        ssHTable.Row = i
        ssHTable.Col = 9
        If svWsCd <> ssHTable.Value Then
            pParas = pParas & ssHTable.Value & "-"
            svWsCd = ssHTable.Value
            ssHTable.Col = 10
            pParas = pParas & ssHTable.Value & ";"
            svWsUnit = ssHTable.Value
        Else
            ssHTable.Col = 10
            If svWsUnit <> ssHTable.Value Then
                ssHTable.Col = 9
                pParas = pParas & ssHTable.Value & "-"
                ssHTable.Col = 10
                pParas = pParas & ssHTable.Value & ";"
                svWsUnit = ssHTable.Value
            End If
        End If
            
    Next
    
    'MyReport.WS2Keys = pParas

    Dim strParas As String, strTmp As String
    Dim strSpcGrp As String, strWsUnit As String
    
    '2000.08.08 추가 : Nogrowth Batch등록에서 보류리스트의 Worksheet을 출력할 경우...
    strParas = pParas
    strTmp = medShift(strParas, ";")
    While (Trim(strTmp) <> "")
        strSpcGrp = medGetP(strTmp, 1, "-")
        strWsUnit = medGetP(strTmp, 2, "-")
        
        Set MyReport = New clsWorkListM
        MyReport.Worksheet2 = True
        Call MyReport.GetInputData(strSpcGrp, strWsUnit, "")
        Call MyReport.PrintReport
        Set MyReport = Nothing
        
        strTmp = medShift(strParas, ";")
    Wend
        
End Sub

Private Sub Form_Load()

    ssTable.Row = 1: ssTable.Col = 1: fGCItem = ssTable.ForeColor

    objMicRst.LoadWorksheetCode MWS_ForCulture, cboWSCode, fWorkSheet
    
    
    cboWSCode.ListIndex = -1: Erase fNGCode
    txtWSUnit = ""
    ssHTable.MaxRows = 0
    ScreenClear

End Sub


Private Sub ScreenClear()

    lblTCount = "": lblRcvDT = "": lblBltDate = ""
    lblCount = "": ssTable.MaxRows = 0
    
End Sub

Private Sub cmdExit_Click()
    
    Unload Me
    Set objRstDic = Nothing
    Set objMicRst = Nothing
    Set objMicCul = Nothing
    Set frm253MReading = Nothing

End Sub

Private Sub cmdWSList_Click()

    
    Dim sWsCd As String

    If cboWSCode.ListIndex < 0 Then Exit Sub

    sWsCd = fWorkSheet(cboWSCode.ListIndex).WsCode
    objMicRst.LoadMicWorkList sWsCd, lstWSUnit
    If lstWSUnit.ListCount <= 0 Then Exit Sub
    
    lstWSUnit.ListIndex = 0
    lstWSUnit.Visible = True
    lstWSUnit.ZOrder
    
End Sub

Private Sub lstWSUnit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim iListIndex As Integer, iWSIndex As Integer
    
    MouseRunning
    
    iWSIndex = cboWSCode.ListIndex
    iListIndex = lstWSUnit.ListIndex
    
    If Button = vbLeftButton And iListIndex >= 0 Then
        txtWSUnit.Text = medGetP(lstWSUnit.List(iListIndex), 1, vbTab)
        lstWSUnit.Clear
        lstWSUnit.Visible = False
        DoEvents
        Call DisplayData(fWorkSheet(iWSIndex).WsCode, txtWSUnit.Text, fWorkSheet(iWSIndex).WsRstType)
    End If
    
    lstWSUnit.Clear
    lstWSUnit.Visible = False
    
    MouseDefault

    
End Sub

Private Sub DisplayData(ByVal pWsCd As String, ByVal pWsUnit As String, ByVal sRTs As String)


    Dim strBuildDtTm As String, strRcvDtTm As String

    ScreenClear

    Call objMicRst.DispWorksheetInfo(pWsCd, pWsUnit, strBuildDtTm, strRcvDtTm)
    lblBltDate.Caption = strBuildDtTm
    lblRcvDT.Caption = strRcvDtTm
    
    lblCount = objMicCul.DispNogrowthList(ssTable, pWsCd, pWsUnit, sRTs)
    
    lblHCount = objMicCul.DispHoldingList(ssHTable, pWsCd, pWsUnit, sRTs, True)


End Sub

Private Sub cmdIn1_Click()
    
    Dim i As Integer, sCnt As Integer
    ssHTable.Col = 1: sCnt = 0
    
    For i = ssHTable.MaxRows To 1 Step -1
        ssHTable.Row = i
        If ssHTable.ForeColor = fSCItem Then
            sCnt = sCnt + 1
            Call AddWorkSheet(ssHTable, i)
        End If
    Next i

    lblHCount = Val(lblHCount) - sCnt
    lblCount = Val(lblCount) + sCnt

End Sub
'
Private Sub ssHTable_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Row < 1 Then Exit Sub

    Call AddWorkSheet(ssHTable, Row)
    lblHCount = Val(lblHCount) - 1
    lblCount = Val(lblCount) + 1

End Sub

Private Sub AddWorkSheet(ByVal pObj As Object, ByVal pRow As Integer)
    
    Dim sAccBuf As String

    ssTable.MaxRows = ssTable.MaxRows + 1
    
    pObj.Col = 1: pObj.COL2 = pObj.MaxCols
    pObj.Row = pRow: pObj.Row2 = pRow
    ssTable.Col = 1: ssTable.COL2 = ssTable.MaxCols
    ssTable.Row = ssTable.MaxRows: ssTable.Row2 = ssTable.MaxRows
    ssTable.Clip = pObj.Clip
    
    Call SaveOneRow(pRow, ssHTable, MWS_Ready)
    
    pObj.Row = pRow
    pObj.Action = ActionDeleteRow
    pObj.MaxRows = pObj.MaxRows - 1
    
End Sub

Private Sub ssTable_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim tmpcolor As Long
    
    If Row = 0 Then
        
        ssTable.Col = -1: ssTable.Row = -1
        ssTable.ForeColor = fGCItem
        
        ssTable.SortBy = SortByRow
        ssTable.SortKey(1) = Col
        ssTable.SortKey(2) = 1
        ssTable.SortKeyOrder(1) = SortKeyOrderAscending
        ssTable.SortKeyOrder(2) = SortKeyOrderAscending
        ssTable.Col = 1
        ssTable.COL2 = ssTable.MaxCols
        ssTable.Row = 1
        ssTable.Row2 = ssTable.MaxRows
        ssTable.Action = ActionSort
        
    End If
    
    If Col >= 0 And Row > 0 Then
    
        ssTable.Col = -1: ssTable.Row = Row
        tmpcolor = ssTable.ForeColor
        
        If tmpcolor = fSCItem Then
            ssTable.ForeColor = fGCItem
        Else
            ssTable.ForeColor = fSCItem
        End If
        
    End If
    
End Sub

Private Sub cmdEx1_Click()
    
    Dim i As Integer, sCnt As Integer
    ssTable.Col = 1: sCnt = 0
    
    For i = ssTable.MaxRows To 1 Step -1
        ssTable.Row = i
        If ssTable.ForeColor = fSCItem Then
            sCnt = sCnt + 1
            MovetoETable i
        End If
    Next i

    lblCount = Val(lblCount) - sCnt
    lblHCount = Val(lblHCount) + sCnt

End Sub

Private Sub ssTable_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Row < 1 Then Exit Sub

    MovetoETable Row
    lblCount = Val(lblCount) - 1
    lblHCount = Val(lblHCount) + 1

End Sub

Private Sub MovetoETable(ByVal pRow As Integer)
    
    Dim sAccBuf As String
    Dim strKey1 As String
    Dim strKey2 As String
    Dim strKey3 As String
    Dim i As Long

    ssHTable.MaxRows = ssHTable.MaxRows + 1
    
    ssTable.Col = 1: ssTable.Row = pRow
    strKey1 = ssTable.Text
    strKey2 = cboWSCode.Text
    strKey3 = txtWSUnit.Text
    
    With ssHTable
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            If .Value = strKey1 Then
                .Col = 9
                If .Value = strKey2 Then
                    .Col = 10
                    If .Value = strKey3 Then Exit Sub
                End If
            End If
        Next
    End With
    
    ssTable.Col = 1: ssTable.COL2 = ssTable.MaxCols
    ssTable.Row = pRow: ssTable.Row2 = pRow
    ssHTable.Col = 1: ssHTable.COL2 = ssTable.MaxCols
    ssHTable.Row = ssHTable.MaxRows: ssHTable.Row2 = ssHTable.MaxRows
    ssHTable.Clip = ssTable.Clip
    ssHTable.Col = 9    '검체군
    ssHTable.Value = fWorkSheet(cboWSCode.ListIndex).WsCode
    ssHTable.Col = 10    'Worksheet Unit
    ssHTable.Value = txtWSUnit.Text
    
    '보류리스트로 옮기는 동시에 Status도 Update한다.
    Call SaveOneRow(pRow, ssTable, MWS_Holding)
    
    ssTable.Row = pRow
    ssTable.Action = ActionDeleteRow
    ssTable.MaxRows = ssTable.MaxRows - 1
    
    With ssHTable
        .ReDraw = False
        .Row = 1: .Row2 = .MaxRows
        .Col = 1: .COL2 = .MaxCols
    
        .SortBy = SortByRow
        .SortKey(1) = 9
        .SortKey(2) = 10
        .SortKey(3) = 1
        .SortKeyOrder(1) = SortKeyOrderAscending
        .SortKeyOrder(2) = SortKeyOrderAscending
        .SortKeyOrder(3) = SortKeyOrderAscending
        .Action = ActionSort
        .ReDraw = True
    End With
    
End Sub


Private Sub SaveOneRow(ByVal iRow As Long, ByVal ssObj As Object, ByVal pStatus As String)
        
    Dim sAccNo As String, sWorkArea As String, sAccDt As String, sAccSeq As String
    Dim sWsCd As String
    
    'pStatus = MWS_Holding - 보류, MWS_Ready - Worksheet
               
    With ssObj
        
        .Col = 1: .Row = iRow: sAccNo = .Text
        sWorkArea = medGetP(sAccNo, 1, "-")
        sAccDt = medGetP(sAccNo, 2, "-")
        sAccSeq = medGetP(sAccNo, 3, "-")
        sAccDt = IIf(Mid$(sAccDt, 1, 1) = "9", "19", "20") & sAccDt
        
        sWsCd = fWorkSheet(cboWSCode.ListIndex).WsCode
        Call objMicCul.SaveOneStatus(sWsCd, txtWSUnit.Text, sWorkArea, sAccDt, sAccSeq, pStatus)
    
    End With

End Sub


Private Sub txtWSUnit_KeyPress(KeyAscii As Integer)
    
    Dim iWSIndex As Integer
        
    If KeyAscii = vbKeyReturn Then
    
        iWSIndex = cboWSCode.ListIndex

        If ExistWS(fWorkSheet(iWSIndex).WsCode, txtWSUnit.Text) Then
            Call DisplayData(fWorkSheet(iWSIndex).WsCode, txtWSUnit.Text, fWorkSheet(iWSIndex).WsRstType)
        Else
            Call ScreenClear
        End If
        
    End If

End Sub

