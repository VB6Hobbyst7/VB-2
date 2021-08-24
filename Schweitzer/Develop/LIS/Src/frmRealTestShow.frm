VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{9167B9A7-D5FA-11D2-86CA-00104BD5476F}#5.0#0"; "DRctl1.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmRealTestShow 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "관련검사조회"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10620
   Icon            =   "frmRealTestShow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H80000005&
      Caption         =   "확인(&O)"
      Height          =   510
      Left            =   9225
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   7140
      Width           =   1320
   End
   Begin DRcontrol1.DrFrame fraLastRst 
      Height          =   7080
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   12488
      Title           =   ""
      TitlePos        =   0
      DelLine         =   0
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComctlLib.ListView lvwLResult 
         Height          =   6630
         Left            =   30
         TabIndex        =   1
         Top             =   390
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   11695
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "접수번호"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "보고일자"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "검사코드"
            Object.Width           =   1411
         EndProperty
      End
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   300
         Index           =   0
         Left            =   45
         TabIndex        =   2
         Top             =   60
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   529
         BackColor       =   8388608
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Alignment       =   1
         Caption         =   "특수검사 관련검사 결과 리스트"
         Appearance      =   0
         LeftGab         =   200
      End
   End
   Begin DRcontrol1.DrFrame DrFrame1 
      Height          =   7095
      Left            =   3420
      TabIndex        =   3
      Top             =   15
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   12515
      Title           =   ""
      TitlePos        =   0
      DelLine         =   0
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   300
         Index           =   2
         Left            =   45
         TabIndex        =   4
         Top             =   60
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   529
         BackColor       =   8388608
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "특수검사 결과"
         Appearance      =   0
         LeftGab         =   200
      End
      Begin RichTextLib.RichTextBox txtLastRst 
         Height          =   6645
         Left            =   30
         TabIndex        =   6
         Top             =   375
         Width           =   7020
         _ExtentX        =   12383
         _ExtentY        =   11721
         _Version        =   393217
         BackColor       =   15658734
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmRealTestShow.frx":000C
      End
   End
   Begin DRcontrol1.DrFrame DrFrame2 
      Height          =   7095
      Left            =   3420
      TabIndex        =   7
      Top             =   15
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   12515
      Title           =   ""
      TitlePos        =   0
      DelLine         =   0
      BackColor       =   14411494
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MedControls1.LisLabel LisLabel7 
         Height          =   300
         Index           =   1
         Left            =   45
         TabIndex        =   8
         Top             =   60
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   529
         BackColor       =   8388608
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         Caption         =   "미생물 관련검사결과"
         Appearance      =   0
         LeftGab         =   200
      End
      Begin FPSpread.vaSpread tblResult 
         Height          =   6630
         Left            =   45
         TabIndex        =   9
         Top             =   375
         Width           =   7005
         _Version        =   196608
         _ExtentX        =   12356
         _ExtentY        =   11695
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         AutoCalc        =   0   'False
         AutoClipboard   =   0   'False
         BackColorStyle  =   3
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridShowHoriz   =   0   'False
         GridShowVert    =   0   'False
         GridSolid       =   0   'False
         MaxCols         =   11
         OperationMode   =   1
         Protect         =   0   'False
         ScrollBars      =   2
         ShadowColor     =   12632256
         ShadowDark      =   12632256
         ShadowText      =   0
         SpreadDesigner  =   "frmRealTestShow.frx":022C
         UnitType        =   0
         UserResize      =   0
         VisibleCols     =   8
         VisibleRows     =   22
         TextTip         =   4
      End
   End
End
Attribute VB_Name = "frmRealTestShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sTestDiv As String

Public Sub SpecialTest(ByVal sPtid As String, ByVal sPtNm As String, ByVal CboRel As ComboBox, ByVal TestDiv As String)
    Dim iTmx        As ListItem
    Dim strTmp      As String
    Dim ii          As Integer
    
    lvwLResult.ListItems.Clear
    
    sTestDiv = TestDiv
    
    For ii = 0 To CboRel.ListCount - 1
        
        With lvwLResult
            If medGetP(CboRel.List(ii), 3, vbTab) = sTestDiv Then
                strTmp = medGetP(CboRel.List(ii), 4, vbTab)
                Set iTmx = .ListItems.Add()
                iTmx.Text = medGetP(strTmp, 1, "-") & "-" & medGetP(strTmp, 2, "-") & "-" & medGetP(strTmp, 3, "-")
                iTmx.SubItems(1) = medGetP(CboRel.List(ii), 5, vbTab)
                iTmx.SubItems(2) = medGetP(strTmp, 4, "-")
               
            End If
        End With
    Next
    Call lvwLResult_ItemClick(iTmx)
    Set iTmx = Nothing
End Sub

Public Sub ComboDisplay(ByVal sTestcd As String, ByVal sCombo As String, ByRef objCombo As Object, _
                        ByRef objSpecial As Object, ByVal objMicro As Object)
    Dim ii As Integer
    Dim aryTmp() As String
    
    
    objSpecial.Visible = False
    objMicro.Visible = False
    
    If P_RealTestMicSpecial = False Then Exit Sub
    
    objCombo.Clear
    
    aryTmp = Split(sCombo, COL_DIV)
    
    For ii = LBound(aryTmp()) To UBound(aryTmp())
        If sTestcd = medGetP(aryTmp(ii), 2, vbTab) Then
            If medGetP(aryTmp(ii), 3, vbTab) = "1" Then objSpecial.Visible = True
            If medGetP(aryTmp(ii), 3, vbTab) = "2" Then objMicro.Visible = True
            
            objCombo.AddItem aryTmp(ii)
        End If
    Next
    If objCombo.ListCount = 0 Then
        objCombo.AddItem "< 없음 >"
    End If
    objCombo.ListIndex = 0

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub



Private Sub Form_Load()
    lvwLResult.ListItems.Clear
    txtLastRst.TextRTF = ""
    medClearTable tblResult
    
End Sub



Private Sub lvwLResult_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Dim sWorkArea   As String
    Dim sAccDt      As String
    Dim sAccSeq     As String
    Dim sTestcd     As String
    Dim strTmp      As String
    
    DoEvents
    txtLastRst.TextRTF = ""
    medClearTable tblResult
    
    strTmp = Item.Text
    
    sWorkArea = medGetP(strTmp, 1, "-")
    sAccDt = medGetP(strTmp, 2, "-")
    sAccSeq = medGetP(strTmp, 3, "-")
    sTestcd = Item.SubItems(2)
    
    If sTestDiv = "1" Then
        Dim objETest    As New clsLISSpecialTest
        LisLabel7(2).Caption = "특수검사결과(" & sWorkArea & "-" & sAccDt & "-" & sAccSeq & " 보고일시 : " & Item.SubItems(1) & " )"
        txtLastRst.TextRTF = objETest.GetResultText(sWorkArea, sAccDt, sAccSeq, sTestcd)
        Set objETest = New clsLISSpecialTest
    Else
        LisLabel7(1).Caption = "미생물관련검사결과(" & sWorkArea & "-" & sAccDt & "-" & sAccSeq & " 보고일시 : " & Item.SubItems(1) & " )"
        Call DisplayMicroResult(sWorkArea, sAccDt, sAccSeq)
        
    End If

End Sub

'% Lab No.를 기준으로 검색한 결과내역을 테이블에 Display한다.
Private Sub DisplayMicroResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As Integer)
   
    
    Dim objResult   As New clsLISResultReview
    Dim i           As Integer
    Dim j           As Integer
    
    With objResult
      
        Call .ResultQuery(pWorkArea, pAccDt, pAccSeq)

        
        ' 일반검사 - High / Low 컬럼 ForeColor 설정
        For i = 1 To .RstRow
            tblResult.Row = i   '+ .OffSet
            For j = 1 To 8
                tblResult.Col = j
                'If .Get_ForeColor(j, i) <> 0 Then tblResult.ForeColor = .Get_ForeColor(j, i)
                tblResult.ForeColor = .Get_ForeColor(j, i)
            Next
        Next
      
        
        '결과내역 Display
        tblResult.Row = 1
        tblResult.Row2 = tblResult.MaxRows
        tblResult.Col = 2
        tblResult.COL2 = tblResult.MaxCols
        tblResult.BlockMode = True
        tblResult.AllowCellOverflow = True
        tblResult.Clip = .ResultClipText '& .SenClipText 'ResultBuffer
        tblResult.BlockMode = False
      
        '미생물 감수성 결과의 경우 항생제명 순으로 Sort / Align Left
        If .SortFg Then
            For i = 1 To .SensiCount
                tblResult.SortBy = SortByRow
                tblResult.SortKey(1) = 2  '항생제명
                tblResult.SortKeyOrder(1) = SortKeyOrderAscending
                tblResult.Col = -1
                tblResult.Row = .AntiSortStartRow(i)   '+ .OffSet
                tblResult.Row2 = .AntiSortEndRow(i)    '+ .OffSet
                tblResult.Action = ActionSort
                tblResult.Row = .SortStartRow - 1 '+ .OffSet
                tblResult.Col = 2
                tblResult.FontUnderline = True
            Next
        Else
            tblResult.Col = 6
            tblResult.Row = -1
            tblResult.ForeColor = DCM_LightRed
            tblResult.FontBold = True
        End If
        If Val(.TestDiv) = TST_MicTest Then
            '미생물 결과 : 균명컬럼 Align Left
            tblResult.Row = -1
            tblResult.Col = -1
            tblResult.BlockMode = True
            tblResult.AllowCellOverflow = True
            tblResult.TypeHAlign = TypeHAlignLeft
            tblResult.BlockMode = False
            tblResult.ColWidth(2) = 17
            'tblResult.ColWidth(3) = 60
            For i = 1 To 5
                If .MicFg(i) Then
                    tblResult.ColWidth(i + 2) = 9
                Else
                    tblResult.ColWidth(i + 2) = 4
                End If
            Next
            tblResult.ColWidth(8) = 20
            tblResult.Col = 3: tblResult.COL2 = 7
            tblResult.Row = -1
            tblResult.BlockMode = True
            tblResult.FontBold = False
            tblResult.BlockMode = False
        Else
            '일반결과 : 결과컬럼 Align Center
            tblResult.Row = 1: tblResult.Row2 = tblResult.MaxRows
            tblResult.Col = 3: tblResult.COL2 = 7
            tblResult.BlockMode = True
            tblResult.TypeHAlign = TypeHAlignCenter
            tblResult.BlockMode = False
            tblResult.ColWidth(2) = 13
            tblResult.ColWidth(3) = 9
            tblResult.ColWidth(4) = 9
            tblResult.ColWidth(5) = 3
            tblResult.ColWidth(6) = 5
            tblResult.ColWidth(7) = 13
        End If
        
        tblResult.Col = 1: tblResult.Row = 1
        
    End With
    
    Set objResult = Nothing
   
End Sub
