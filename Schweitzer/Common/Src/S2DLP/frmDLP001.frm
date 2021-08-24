VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDLP001 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "코드 찾기"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   Icon            =   "frmDLP001.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame fraDLP 
      BackColor       =   &H00EEEEEE&
      Caption         =   "검색"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15
      TabIndex        =   0
      Top             =   3105
      Width           =   3470
      Begin VB.OptionButton optCode 
         BackColor       =   &H00EEEEEE&
         Caption         =   "코드명"
         Height          =   180
         Index           =   1
         Left            =   750
         TabIndex        =   3
         Top             =   285
         Width           =   840
      End
      Begin VB.OptionButton optCode 
         BackColor       =   &H00EEEEEE&
         Caption         =   "코드"
         Height          =   180
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Top             =   270
         Value           =   -1  'True
         Width           =   660
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFBFF&
         Height          =   330
         Left            =   1635
         ScrollBars      =   2  '수직
         TabIndex        =   1
         Top             =   210
         Width           =   1710
      End
   End
   Begin MSComctlLib.ListView lvwCodeList 
      Height          =   3030
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   5345
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16776191
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "코드"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "코드명"
         Object.Width           =   3175
      EndProperty
   End
   Begin MSComctlLib.ListView lvwSort 
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1508
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Key"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Data"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmDLP001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mvarSqlStmt As String                       'SQL 문장을 받는 변수
Private mvarDicObj As clsDictionary                 '디스플레이할 Dictionary Object
Private blnFirst As Boolean
Private mvarHeadName As String                      '리스트뷰의 컬럼헤더명
Private mvarColSize As New Scripting.Dictionary     '리스트뷰의 컬럼헤더 사이즈
Private objProbar As clsProgressBar

'
Public Event ListSelected(ByVal strSelList As String)   '리스트뷰에서 값을 넘겨받는 이벤트

'[속성] - Sql문장
Public Property Let SqlStmt(ByVal vData As String)
    mvarSqlStmt = vData
End Property

Public Property Let HeadName(ByVal vData As String)
    mvarHeadName = vData
End Property

Public Property Set DicObj(ByVal vData As clsDictionary)
    Set mvarDicObj = vData
End Property

Public Property Set ColSize(ByVal vData As Object)
    Set mvarColSize = vData
End Property

Private Sub Form_Activate()
    '
    If Not (mvarDicObj Is Nothing) Then
        Call CodeListDicView
    End If
    '
    DoEvents
    If mvarSqlStmt <> "" Then
        Call CodeListView(lvwCodeList, mvarSqlStmt, mvarHeadName)
'        DoEvents
        blnFirst = True
    End If
    
'Modify By Legends -- 2001/03/29
'    DoEvents
    '
'    If blnFirst = False And mvarSqlStmt <> "" Then
'        Call CodeListView(lvwCodeList, mvarSqlStmt, mvarHeadName)
''        DoEvents
'        blnFirst = True
'    End If
''    DoEvents
'    '
End Sub

Private Sub Form_Load()
    
'Modify By Legends -- 2001/03/29
    
'    If mvarSqlStmt <> "" Then
'        Call CodeListView(lvwCodeList, mvarSqlStmt, mvarHeadName)
''        DoEvents
'        blnFirst = True
'    End If
    '
End Sub

Private Sub Form_Terminate()
    '
    mvarColSize.RemoveAll
    Set mvarDicObj = Nothing
    blnFirst = False
    mvarHeadName = ""
    mvarSqlStmt = ""
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mvarColSize.RemoveAll
    Set mvarDicObj = Nothing
    blnFirst = False
    mvarHeadName = ""
    mvarSqlStmt = ""
    
End Sub

Private Sub lvwCodeList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'소트
    Static lngOrder(1) As Integer
    
    With lvwCodeList
'        .SortKey = ColumnHeader.index - 1
'        .SortOrder = IIf(intOrder = 0, lvwAscending, lvwDescending)
'        .Sorted = True
'        intOrder = (intOrder + 1) Mod 2
        Select Case ColumnHeader.index - 1
            Case 0
                .SortKey = 0
                .SortOrder = IIf(lngOrder(0) = 0, lvwAscending, lvwDescending)
                .Sorted = True
                lngOrder(0) = (lngOrder(0) + 1) Mod 2
            Case 1
                .SortKey = 1
                .SortOrder = IIf(lngOrder(1) = 0, lvwAscending, lvwDescending)
                .Sorted = True
                lngOrder(1) = (lngOrder(1) + 1) Mod 2
        End Select
    End With
End Sub

Private Sub lvwCodeList_DblClick()
    Dim strSelText As String
    Dim ItemX As ListItem
    Dim i As Long
    
    If lvwCodeList.ListItems.Count < 1 Then Exit Sub
    '
    Set ItemX = lvwCodeList.SelectedItem
    strSelText = ItemX.Text
    For i = 1 To lvwCodeList.ColumnHeaders.Count - 1
        strSelText = strSelText & ";" & ItemX.SubItems(i)
    Next
    RaiseEvent ListSelected(strSelText)
    
End Sub


Private Sub optCode_Click(index As Integer)
    txtSearch.SetFocus
End Sub

Private Sub txtSearch_Change()
'찾기
    Dim strFindItem As String
    Dim itmFound As ListItem   ' FoundItem 변수입니다.
    Dim itmX As ListItem
    Dim i As Long
        
    strFindItem = Trim(txtSearch)
    
    With lvwCodeList
        If optCode(0).Value Then
            For i = 1 To .ListItems.Count
                Set itmX = .ListItems(i)
                If UCase(itmX.Text) Like UCase(strFindItem & "*") Then
                    itmX.Selected = True
                    itmX.EnsureVisible
                    Exit For
'                Else
'                        With txtSearch
'                            .SelStart = 0
'                            .SelLength = Len(strFindItem)
'                        End With
                End If
            Next
        Else
            For i = 1 To .ListItems.Count
                Set itmX = .ListItems(i)
                If (itmX.SubItems(1) Like (strFindItem & "*")) Then
                    itmX.Selected = True
                    itmX.EnsureVisible
                    Exit For
                Else
'                    With txtSearch
'                        .SelStart = 0
'                        .SelLength = Len(strFindItem)
'                    End With
                End If
            Next
        End If
    End With

End Sub


Private Function CodeListView(ByRef pListView As Object, ByVal strSQL As String, _
    ByVal pHeadName As String)
'리스트 뷰에 보여주는
'리스트뷰를 받는 파라미터
'SQL문장을 받는 파라미터

    Dim Rs As New RecordSet
    Dim itmX As Object
    Dim i As Long
    
    Set objProbar = New clsProgressBar
    
    With objProbar
        .SetMyForm Me
        .XPos = 0  'Me.ScaleHeight - (Me.ScaleHeight - 300)
        .YHeight = 280
        .Choice = True
        .Msg = "자료를 읽기 위해 준비중입니다..."
        .Value = 1
    End With
    
    Rs.Open strSQL, DBConn
'    If rs.DBerror Then
'        DisplayErrors
'        Exit Function
'    End If

    With pListView
        If pHeadName <> "" Then
            Dim aryTmp() As String
            Dim ii As Long
            aryTmp = Split(pHeadName, ",")
            ReDim Preserve aryTmp(1)
            For ii = LBound(aryTmp) To UBound(aryTmp)
                If mvarColSize.Exists(ii + 1) = True Then
                    aryTmp(ii) = mvarColSize(ii + 1)
                End If
            Next ii
            medInitLvwHead pListView, Join(aryTmp, ","), "-700,400"
        End If
        .ListItems.Clear
'        If Rs.DBerror = True Then
'            MsgBox "오류가 발생하였읍니다", vbCritical, "오류"
'        Else
            objProbar.Max = Rs.RecordCount
            objProbar.Msg = ""
            Do Until Rs.EOF
                i = i + 1
                Set itmX = .ListItems.Add(, , "" & Rs.Fields(0).Value)
                itmX.SubItems(1) = "" & Rs.Fields(1).Value
                Rs.MoveNext
                objProbar.Value = i
            Loop
'        End If
    End With
'    DoEvents

On Error Resume Next
    Rs.Close
    Set Rs = Nothing
    Set objProbar = Nothing
End Function

Private Sub CodeListDicView()
Dim aryTmp() As String
Dim ii As Long


    If mvarSqlStmt = "" Then
        '
        If mvarDicObj Is Nothing Then Exit Sub
        
        With mvarDicObj
            If .RecordCount > 0 Then
                lvwCodeList.ListItems.Clear
                If mvarHeadName <> "" Then
                    aryTmp = Split(mvarHeadName, ",")
                    ReDim Preserve aryTmp(.ColCount - 1)
                Else
                    aryTmp = Split(.FieldName, ",")
                End If
                '
                If UBound(aryTmp) = 1 Then
                    aryTmp(0) = -700: aryTmp(1) = "400"
                ElseIf UBound(aryTmp) = 2 Then
                    aryTmp(0) = "-300"
                    aryTmp(1) = "200"
                    aryTmp(2) = "-200"
                Else
                    For ii = LBound(aryTmp) To UBound(aryTmp)
                        aryTmp(ii) = "0"
                        If ii = 0 Then
                            aryTmp(ii) = "-100"
                        ElseIf ii = 1 Then
                            aryTmp(ii) = "100"
                        End If
                    Next ii
                End If
                For ii = LBound(aryTmp) To UBound(aryTmp)
                    If mvarColSize.Exists(CStr(ii + 1)) = True Then
                        aryTmp(ii) = mvarColSize(CStr(ii + 1))
                    End If
                Next ii
                If mvarHeadName <> "" Then
                    medInitLvwHead lvwCodeList, mvarHeadName, Join(aryTmp, ",")
                Else
                    medInitLvwHead lvwCodeList, .FieldName, Join(aryTmp, ",")
                End If
                '
                medDataLoadLvw lvwCodeList, LINE_DIV, COL_DIV, .GetString, .GetTagString
'                DoEvents
            End If
        End With
        '
        Exit Sub
        '
    End If
    '
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtSearch.Text) <> "" Then
            Call lvwCodeList_DblClick
        End If
    End If
End Sub
