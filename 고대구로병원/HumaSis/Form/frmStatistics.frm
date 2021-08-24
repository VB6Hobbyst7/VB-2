VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4BD5DFC7-B668-44E0-A002-C1347061239D}#1.0#0"; "HSCotrol.ocx"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmStatistics 
   Caption         =   "°á°ú ÀÔ·Â"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   11970
   WindowState     =   2  'ÃÖ´ëÈ­
   Begin BHButton.BHImageButton cmdSerch 
      Height          =   330
      Left            =   4950
      TabIndex        =   13
      Top             =   675
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      Caption         =   "Á¶È¸"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin VB.OptionButton optCondition 
      Caption         =   "½½¸³ °Ç¼ö"
      Height          =   285
      Index           =   1
      Left            =   9915
      TabIndex        =   8
      Top             =   675
      Width           =   1215
   End
   Begin VB.OptionButton optCondition 
      Caption         =   "°Ë»ç °Ç¼ö"
      Height          =   285
      Index           =   0
      Left            =   8475
      TabIndex        =   7
      Top             =   675
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwStatics 
      Height          =   5415
      Index           =   0
      Left            =   30
      TabIndex        =   5
      Top             =   1035
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   9551
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imlList 
      Left            =   10770
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatistics.frx":0000
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatistics.frx":059A
            Key             =   "ERR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatistics.frx":0B34
            Key             =   "NOF"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatistics.frx":10CE
            Key             =   "LST"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatistics.frx":1668
            Key             =   "LSE"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatistics.frx":1C02
            Key             =   "LSN"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraCmdBar 
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   1.5
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   15
      TabIndex        =   1
      Top             =   6465
      Width           =   11940
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Save"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   1
         Left            =   1410
         TabIndex        =   10
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Delete"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   2
         Left            =   2730
         TabIndex        =   11
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Clear"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdAction 
         Height          =   420
         Index           =   3
         Left            =   4050
         TabIndex        =   12
         Top             =   90
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   741
         Caption         =   "Close"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
   End
   Begin HSCotrol.CaptionBar CaptionBar1 
      Align           =   1  'À§ ¸ÂÃã
      Height          =   555
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11970
      _ExtentX        =   21114
      _ExtentY        =   979
      Border          =   1
      CaptionBackColor=   16777215
      Picture         =   "frmStatistics.frx":219C
      Caption         =   " Test Statistics"
      SubCaption      =   "°Ë»ç Åë°è."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SubCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpFromDate 
      Height          =   300
      Left            =   1410
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   690
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   68026369
      CurrentDate     =   37112
   End
   Begin MSComCtl2.DTPicker dtpToDate 
      Height          =   300
      Left            =   3180
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   690
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   68026369
      CurrentDate     =   37112
   End
   Begin MSComctlLib.ListView lvwStatics 
      Height          =   5415
      Index           =   1
      Left            =   30
      TabIndex        =   6
      Top             =   1035
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   9551
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Á¢¼öÀÏ :"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   555
      TabIndex        =   4
      Top             =   750
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   330
      Left            =   165
      Picture         =   "frmStatistics.frx":341E
      Stretch         =   -1  'True
      Top             =   675
      Width           =   330
   End
End
Attribute VB_Name = "frmStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mAdoRs As ADODB.Recordset

Private Const COL_WIDTH As Long = "900"

Private Sub cmdAction_Click(Index As Integer)
    Select Case Index
        Case 0
            Call cmdSave_Click
        Case 1
            Call cmdPrint2_Click
        Case 2
            Call cmdClear_Click
        Case 3 'cmd close
            Call cmdClose_Click
        Case Else
    End Select
End Sub

Private Sub cmdSave_Click()
End Sub

Private Sub cmdPrint2_Click()

End Sub

Private Sub cmdClear_Click()
    With lvwStatics(0)
        .HideColumnHeaders = True
        .ColumnHeaders.Clear
        .ListItems.Clear
    End With
    With lvwStatics(1)
        .HideColumnHeaders = True
        .ColumnHeaders.Clear
        .ListItems.Clear
    End With
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Progress_View()
Dim I As Long

    Call SetProgress(1005, Custom, "Loding", True)
    
    For I = 1 To 1001
        Call ShowProgress(I, "TEST " & I, True)
    Next
    Call SetProgress(100, Custom, "End", False)
End Sub

Private Sub cmdSerch_Click()
    
    With lvwStatics(0)
        .HideColumnHeaders = True
        .ColumnHeaders.Clear
        .ListItems.Clear
    End With
    
    If Not Drow_Header Then
        Exit Sub
    End If
    
    If Not Drow_Date(0) Then
        Exit Sub
    End If
    
    If Not Drow_Item Then
        Exit Sub
    End If
    Call Total_Calculation(0)
    
    With lvwStatics(1)
        .HideColumnHeaders = True
        .ColumnHeaders.Clear
        .ListItems.Clear
        .ColumnHeaders.Add , "KEY_DATE", "DATE"
        .ColumnHeaders.Add , "KEY_SLIP", "SLIP COUNT", , lvwColumnRight
        .ColumnHeaders.Add , "KEY_TEST", "TEST COUNT", , lvwColumnRight
        .HideColumnHeaders = False
    End With
    
    If Not Drow_Date(1) Then
        Exit Sub
    End If
    If Not Drow_SlipCount Then
        Exit Sub
    End If
    If Not Drow_TestCount Then
        Exit Sub
    End If
    Call Total_Calculation(1)

End Sub

Private Function Drow_Header() As Boolean
    Dim objStatics      As clsStatistics
    Dim itemKey         As String
    Dim itemText        As String
    Dim AdoRs_TstNm     As ADODB.Recordset
    
    Set objStatics = New clsStatistics
    Drow_Header = True
    With objStatics
        .SetAdoCn AdoCn_Jet
        
        Set AdoRs_TstNm = .Get_TestName(Format(dtpFromDate, "YYYY/MM/DD"), Format(dtpToDate, "YYYY/MM/DD"))
        If Not AdoRs_TstNm Is Nothing Then
            If AdoRs_TstNm.EOF Then
                Drow_Header = False
            Else
                lvwStatics(0).ColumnHeaders.Clear
                lvwStatics(0).ColumnHeaders.Add , "DATE", "DATE"
                AdoRs_TstNm.MoveFirst
                Do Until AdoRs_TstNm.EOF
                    itemKey = Trim(AdoRs_TstNm.Fields("TESTCD") & "")
                    itemText = Trim(AdoRs_TstNm.Fields("TESTNM") & "")
                    Call lvwStatics(0).ColumnHeaders.Add(, itemKey, itemText, COL_WIDTH, lvwColumnRight)
                    AdoRs_TstNm.MoveNext
                Loop
                lvwStatics(0).HideColumnHeaders = False
            End If
        Else
            Drow_Header = False
        End If
    End With
    
    Set AdoRs_TstNm = Nothing
    Set objStatics = Nothing
End Function

Private Function Drow_Date(ByVal Index As Integer) As Boolean
    Dim objStatics      As clsStatistics
    Dim itemKey         As String
    Dim itemText        As String
    Dim AdoRs_TstDt     As ADODB.Recordset
    
    Set objStatics = New clsStatistics
    Drow_Date = True
    With objStatics
        .SetAdoCn AdoCn_Jet
        Set AdoRs_TstDt = .Get_TestDate(Format(dtpFromDate, "YYYY/MM/DD"), Format(dtpToDate, "YYYY/MM/DD"))
        If Not AdoRs_TstDt Is Nothing Then
            If AdoRs_TstDt.EOF Then
                Drow_Date = False
            Else
                AdoRs_TstDt.MoveFirst
                Do Until AdoRs_TstDt.EOF
                    itemKey = Trim(AdoRs_TstDt.Fields("ACCDT") & "")
                    itemText = Trim(AdoRs_TstDt.Fields("ACCDT") & "")
                    lvwStatics(Index).ListItems.Add , itemKey, itemText, , "LST"
                    AdoRs_TstDt.MoveNext
                Loop
            End If
        Else
            Drow_Date = False
        End If
    End With
    
    Set AdoRs_TstDt = Nothing
    Set objStatics = Nothing
End Function

Private Function Drow_Item() As Boolean
    Dim objStatics      As clsStatistics
    Dim itemKey         As String
    Dim itemHeadKey     As String
    Dim itemText        As String
    Dim AdoRs_TstCn     As ADODB.Recordset
    
    Set objStatics = New clsStatistics
    Drow_Item = True
    With objStatics
        .SetAdoCn AdoCn_Jet
        Set AdoRs_TstCn = .Get_TestCount(Format(dtpFromDate, "YYYY/MM/DD"), Format(dtpToDate, "YYYY/MM/DD"))
        If Not AdoRs_TstCn Is Nothing Then
            If AdoRs_TstCn.EOF Then
                Drow_Item = False
            Else
                Do Until AdoRs_TstCn.EOF
                    itemKey = Trim(AdoRs_TstCn.Fields("ACCDT") & "")
                    itemHeadKey = Trim(AdoRs_TstCn.Fields("TESTCD") & "")
                    itemText = Trim(AdoRs_TstCn.Fields("CNT") & "")
                    lvwStatics(0).ListItems(itemKey).SubItems(lvwStatics(0).ColumnHeaders(itemHeadKey).SubItemIndex) = itemText
                    AdoRs_TstCn.MoveNext
                Loop
            End If
        Else
            Drow_Item = False
        End If
        
    End With
    Set AdoRs_TstCn = Nothing
    Set objStatics = Nothing
End Function

Private Sub Total_Calculation(ByVal Index As Integer)
    Dim itemX           As ListItem
    Dim itemS           As ListSubItem
    Dim lngTotal()      As Long
    Dim I As Long
    
    ReDim lngTotal(lvwStatics(Index).ColumnHeaders.Count - 1)
    For Each itemX In lvwStatics(Index).ListItems
        For I = 1 To lvwStatics(Index).ColumnHeaders.Count - 1
            lngTotal(I) = lngTotal(I) + Val(itemX.SubItems(I))
        Next
    Next
    Set itemX = lvwStatics(Index).ListItems.Add
    With itemX
        .text = "TOTAL"
        .Bold = True
    End With
    
    For I = 1 To lvwStatics(Index).ColumnHeaders.Count - 1
        Set itemS = itemX.ListSubItems.Add(I)
        With itemS
            .Bold = True
            .ForeColor = vbBlue
            .text = lngTotal(I)
        End With
    Next
    
    Set itemX = Nothing
    
End Sub

Private Function Drow_SlipCount() As Boolean
    Dim objStatics      As clsStatistics
    Dim itemKey         As String
    Dim itemHeadKey     As String
    Dim itemText        As String
    Dim AdoRs_TstCn     As ADODB.Recordset
    
    Set objStatics = New clsStatistics
    Drow_SlipCount = True
    With objStatics
        .SetAdoCn AdoCn_Jet
        Set AdoRs_TstCn = .Get_SlipCount(Format(dtpFromDate, "YYYY/MM/DD"), Format(dtpToDate, "YYYY/MM/DD"))
        If Not AdoRs_TstCn Is Nothing Then
            If AdoRs_TstCn.EOF Then
                Drow_SlipCount = False
            Else
                Do Until AdoRs_TstCn.EOF
                    itemKey = Trim(AdoRs_TstCn.Fields("ACCDT") & "")
                    itemHeadKey = "KEY_SLIP"
                    itemText = Trim(AdoRs_TstCn.Fields("SLIP_CNT") & "")
                    lvwStatics(1).ListItems(itemKey).SubItems(lvwStatics(1).ColumnHeaders(itemHeadKey).SubItemIndex) = itemText
                    AdoRs_TstCn.MoveNext
                Loop
            End If
        Else
            Drow_SlipCount = False
        End If
        
    End With
    Set AdoRs_TstCn = Nothing
    Set objStatics = Nothing
End Function

Private Function Drow_TestCount() As Boolean
    Dim objStatics      As clsStatistics
    Dim itemKey         As String
    Dim itemHeadKey     As String
    Dim itemText        As String
    Dim AdoRs_TstCn     As ADODB.Recordset
    
    Set objStatics = New clsStatistics
    Drow_TestCount = True
    With objStatics
        .SetAdoCn AdoCn_Jet
        Set AdoRs_TstCn = .Get_TotalTestCount(Format(dtpFromDate, "YYYY/MM/DD"), Format(dtpToDate, "YYYY/MM/DD"))
        If Not AdoRs_TstCn Is Nothing Then
            If AdoRs_TstCn.EOF Then
                Drow_TestCount = False
            Else
                Do Until AdoRs_TstCn.EOF
                    itemKey = Trim(AdoRs_TstCn.Fields("ACCDT") & "")
                    itemHeadKey = "KEY_TEST"
                    itemText = Trim(AdoRs_TstCn.Fields("TEST_CNT") & "")
                    lvwStatics(1).ListItems(itemKey).SubItems(lvwStatics(1).ColumnHeaders(itemHeadKey).SubItemIndex) = itemText
                    AdoRs_TstCn.MoveNext
                Loop
            End If
        Else
            Drow_TestCount = False
        End If
        
    End With
    Set AdoRs_TstCn = Nothing
    Set objStatics = Nothing
End Function

Private Sub Form_Load()
    Dim itemX As ListItem

    With lvwStatics(0)
        .View = lvwReport
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .ColumnHeaderIcons = imlList
        .SmallIcons = imlList
        .HideColumnHeaders = True
    End With
    
    With lvwStatics(1)
        .View = lvwReport
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .ColumnHeaderIcons = imlList
        .SmallIcons = imlList
        .HideColumnHeaders = True
    End With
    
    dtpFromDate.Value = Now
    dtpToDate.Value = Now
    optCondition(0).Value = True
End Sub

Private Sub Form_Resize()
    Dim I As Integer
    If ScaleHeight < 650 Then Exit Sub
    If ScaleWidth < 60 Then Exit Sub
    
    Call lvwStatics(1).Move(lvwStatics(0).left, lvwStatics(0).Top, lvwStatics(0).Width, lvwStatics(0).Height)
    Call fraCmdBar.Move(ScaleLeft + 30, ScaleHeight - fraCmdBar.Height - 30, ScaleWidth - 60)
    
    For I = cmdAction.LBound To cmdAction.UBound
        Call cmdAction(I).Move(fraCmdBar.Width - ((1300 * (cmdAction.Count - I)) + (70 * (cmdAction.UBound - I)) + 100), _
                               (fraCmdBar.Height - 360) / 2, 1300, 360)
    Next
    
End Sub

Private Sub optCondition_Click(Index As Integer)
    lvwStatics(Index).ZOrder
End Sub
