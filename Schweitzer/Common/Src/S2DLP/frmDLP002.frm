VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDLP002 
   BackColor       =   &H00EEEEEE&
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3495
   Icon            =   "frmDLP002.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   3495
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
      Left            =   0
      TabIndex        =   0
      Top             =   3120
      Width           =   3495
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFBFF&
         Height          =   330
         Left            =   60
         ScrollBars      =   2  '수직
         TabIndex        =   1
         ToolTipText     =   "검색할 내용을 입력하십시오."
         Top             =   210
         Width           =   3375
      End
   End
   Begin MSComctlLib.ListView lvwCodeList 
      Height          =   3045
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Double-Click로 선택하거나 선택후 Enter키를 치십시오."
      Top             =   0
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   5371
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
End
Attribute VB_Name = "frmDLP002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Coding By legends 2003/11/21

'컬럼 갯수의 제한이 없다.
'폼의 크기를 외부에서 조정할 수 있다.
'헤더의 폭을 조정할 수 있다.
'외부에서 레코셋 개체를 넘겨받을 수 있다.

Public Event SelectedItem(ByVal pSelectedItem As String)

Private mvarColumnHeaderText As String        '헤더명
Private mvarColumnHeaderWidth As String       '컬럼크기(기본값은 1440)
Private mvarColumnHeaderAlign As String       '컬럼정렬(기본값은 왼쪽 정렬)

Private mvarReturnColumnHeaderWidth As String
Private mvarReturnFormWidth As Long
Private mvarReturnFormHeight As Long
Private mvarReturnFormTop As Long
Private mvarReturnFormLeft As Long

Private mvarSqlStmt As String       '쿼리
Private mvarRecordSet As Object     '레코드셋 개체
Private mvarDivision As String
Private mvarClick As Boolean

'Public Property Let FormHeight(ByVal vData As Long)
'    mvarFormHeight = vData
'End Property
'
'Public Property Get FormHeight() As Long
'    FormHeight = mvarFormHeight
'End Property
'
'Public Property Let FormWidth(ByVal vData As Long)
'    mvarFormWidth = vData
'End Property
'
'Public Property Get FormWidth() As Long
'    FormWidth = mvarFormWidth
'End Property
Public Property Let ColumnHeaderText(ByVal vData As String)
    mvarColumnHeaderText = vData
End Property

Public Property Get ColumnHeaderText() As String
    ColumnHeaderText = mvarColumnHeaderText
End Property

Public Property Let ColumnHeaderWidth(ByVal vData As String)
    mvarColumnHeaderWidth = vData
End Property

Public Property Get ColumnHeaderWidth() As String
    ColumnHeaderWidth = mvarColumnHeaderWidth
End Property

Public Property Let ColumnHeaderAlign(ByVal vData As String)
    mvarColumnHeaderAlign = vData
End Property

Public Property Get ColumnHeaderAlign() As String
    ColumnHeaderAlign = mvarColumnHeaderAlign
End Property

Private Property Let ReturnColumnHeaderWidth(ByVal vData As String)
'읽기 전용
    mvarReturnColumnHeaderWidth = vData
End Property

Public Property Get ReturnColumnHeaderWidth() As String
    ReturnColumnHeaderWidth = mvarReturnColumnHeaderWidth
End Property

Private Property Let ReturnFormWidth(ByVal vData As Long)
    mvarReturnFormWidth = vData
End Property

Public Property Get ReturnFormWidth() As Long
    ReturnFormWidth = mvarReturnFormWidth
End Property

Private Property Let ReturnFormHeight(ByVal vData As Long)
    mvarReturnFormHeight = vData
End Property

Public Property Get ReturnFormHeight() As Long
    ReturnFormHeight = mvarReturnFormHeight
End Property

Private Property Let ReturnFormTop(ByVal vData As Long)
    mvarReturnFormTop = vData
End Property

Public Property Get ReturnFormTop() As Long
    ReturnFormTop = mvarReturnFormTop
End Property

Private Property Let ReturnFormLeft(ByVal vData As Long)
    mvarReturnFormLeft = vData
End Property

Public Property Get ReturnFormLeft() As Long
    ReturnFormLeft = mvarReturnFormLeft
End Property

Public Property Let SqlStmt(ByVal vData As String)
    mvarSqlStmt = vData
End Property

Public Property Get SqlStmt() As String
    SqlStmt = mvarSqlStmt
End Property

Public Property Set RecordSet(ByVal vData As Object)
    Set mvarRecordSet = vData
End Property

Public Property Get RecordSet() As Object
    Set RecordSet = mvarRecordSet
End Property

Public Property Let Division(ByVal vData As String)
    mvarDivision = IIf(vData = "", ";", vData)
End Property

Public Property Get Division() As String
    Division = IIf(mvarDivision = "", ";", mvarDivision)
End Property

Public Property Let Click(ByVal vData As Boolean)
    mvarClick = vData
End Property

Public Property Get Click() As Boolean
    Click = mvarClick
End Property

Private Function LoadList()
'리스트 뷰에 보여주는
'리스트뷰를 받는 파라미터
'SQL문장을 받는 파라미터
    
    Dim objPro As clsProgressBar
    Dim Rs As RecordSet
    Dim itmX As ListItem
    Dim headX As ColumnHeader
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Set objPro = New clsProgressBar
    
    With objPro
        .SetMyForm Me
        .XPos = 0  'Me.ScaleHeight - (Me.ScaleHeight - 300)
        .YHeight = 280
        .Choice = True
        .Msg = "자료를 읽기 위해 준비중입니다..."
        .Value = 1
    End With
    
    On Error GoTo NoData
    
    If mvarSqlStmt = "" And mvarRecordSet Is Nothing Then
        MsgBox "SQL 구문이나 RecordSet 개체를 Property로 지정해야 합니다.", vbCritical
        GoTo NoData
    ElseIf mvarSqlStmt = "" And Not (mvarRecordSet Is Nothing) Then
        Set Rs = mvarRecordSet
    ElseIf mvarSqlStmt <> "" And mvarRecordSet Is Nothing Then
        Set Rs = New RecordSet
        
        Rs.Open mvarSqlStmt, DBConn
    ElseIf mvarSqlStmt <> "" And Not (mvarRecordSet Is Nothing) Then    '둘다 넘어왔으면 레코드셋 우선
        Set Rs = mvarRecordSet
    End If
    
    With lvwCodeList
        If mvarColumnHeaderText <> "" Then    'ColumnHeader 셋팅
            
            Dim aryHeader() As String
            
            aryHeader = Split(mvarColumnHeaderText, mvarDivision)
            
            .ColumnHeaders.Clear
            If UBound(aryHeader) < 0 Then  '헤더가 제대로 안넘어 온경우
                Set headX = .ColumnHeaders.Add()
                headX.Text = "코드"
                
                Set headX = .ColumnHeaders.Add()
                headX.Text = "코드명"
            Else
                For i = LBound(aryHeader) To UBound(aryHeader)
                    Set headX = .ColumnHeaders.Add()
                    headX.Text = aryHeader(i)
                    headX.Width = Val(IIf(medGetP(mvarColumnHeaderWidth, i + 1, mvarDivision) = "", 1440, medGetP(mvarColumnHeaderWidth, i + 1, mvarDivision)))
                    headX.Alignment = IIf(medGetP(mvarColumnHeaderAlign, i + 1, mvarDivision) = "", lvwColumnLeft, medGetP(mvarColumnHeaderAlign, i + 1, mvarDivision))
                Next
            End If
        End If
        
        If Rs.EOF Then GoTo NoData
        
        objPro.Max = Rs.RecordCount
        objPro.Msg = ""
        
        Do Until Rs.EOF
            Set itmX = .ListItems.Add()
            itmX.Text = Rs.Fields(0).Value & ""
            For j = 1 To .ColumnHeaders.Count - 1
                On Error Resume Next
                itmX.SubItems(j) = Rs.Fields(j).Value & ""
            Next
            
            k = k + 1: objPro.Value = k
            
            Rs.MoveNext
        Loop
    End With
    
NoData:
    On Error Resume Next
    Rs.Close
    Set Rs = Nothing
    Set objPro = Nothing
    
    If lvwCodeList.ListItems.Count <> 0 Then
        Call lvwCodeList_ColumnClick(lvwCodeList.ColumnHeaders(1))
    End If
    
End Function

Private Sub Form_Activate()
    Static blnFirst As Boolean

    If blnFirst = False Then Call LoadList
    blnFirst = True
End Sub

Private Sub Form_Initialize()
'    mvarFormWidth = 0
'    mvarFormHeight = 0
    
    mvarColumnHeaderText = ""       '헤더명
    mvarColumnHeaderWidth = ""      '컬럼크기(기본값은 1440)
    mvarColumnHeaderAlign = ""      '컬럼정렬(기본값은 왼쪽 정렬)
    
    mvarReturnColumnHeaderWidth = ""
    mvarReturnFormWidth = 0
    mvarReturnFormHeight = 0
    mvarReturnFormTop = 0
    mvarReturnFormLeft = 0
    
    mvarSqlStmt = ""      '쿼리
    Set mvarRecordSet = Nothing    '레코드셋 개체
    mvarDivision = ""
    mvarClick = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mvarColumnHeaderText = ""       '헤더명
    mvarColumnHeaderWidth = ""      '컬럼크기(기본값은 1440)
    mvarColumnHeaderAlign = ""      '컬럼정렬(기본값은 왼쪽 정렬)
    
    mvarReturnColumnHeaderWidth = ""
    mvarReturnFormWidth = 0
    mvarReturnFormHeight = 0
    mvarReturnFormTop = 0
    mvarReturnFormLeft = 0
    
    mvarSqlStmt = ""      '쿼리
    Set mvarRecordSet = Nothing    '레코드셋 개체
    mvarDivision = ""
    mvarClick = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Screen.ActiveForm.Name <> frmDLP002.Name Then Exit Sub
    
    With frmDLP002
'        If .Width > mvarFormWidth Then
            .lvwCodeList.Width = .ScaleWidth
            .fraDLP.Width = .lvwCodeList.Width
            .txtSearch.Width = .fraDLP.Width - 150
'        End If
        '
'        If .Height > mvarFormHeight Then
            .lvwCodeList.Height = IIf(.fraDLP.Visible, .ScaleHeight - (.fraDLP.Height + 75), .ScaleHeight)
            .fraDLP.Top = .lvwCodeList.Top + .lvwCodeList.Height + 75
'        End If
    End With
End Sub

Private Sub Form_Terminate()
    mvarColumnHeaderText = ""       '헤더명
    mvarColumnHeaderWidth = ""      '컬럼크기(기본값은 1440)
    mvarColumnHeaderAlign = ""      '컬럼정렬(기본값은 왼쪽 정렬)
    
    mvarReturnColumnHeaderWidth = ""
    mvarReturnFormWidth = 0
    mvarReturnFormHeight = 0
    mvarReturnFormTop = 0
    mvarReturnFormLeft = 0
    
    mvarSqlStmt = ""      '쿼리
    Set mvarRecordSet = Nothing    '레코드셋 개체
    mvarDivision = ""
    mvarClick = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmDLP002 = Nothing
End Sub

Private Sub lvwCodeList_Click()
    On Error Resume Next
    If Screen.ActiveControl.Name <> lvwCodeList.Name Then Exit Sub
    
    If mvarClick Then Call lvwCodeList_DblClick
End Sub

Private Sub lvwCodeList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'소트
    Static blnToggle() As Boolean
    Static blnFirst As Boolean
    Dim i As Long
    
    If blnFirst = False Then
        ReDim blnToggle(lvwCodeList.ColumnHeaders.Count - 1)
        blnFirst = True
    End If
    
    '▲▼
    
    For i = 1 To lvwCodeList.ColumnHeaders.Count
        lvwCodeList.ColumnHeaders(i).Text = Trim(Replace(lvwCodeList.ColumnHeaders(i).Text, "▲", ""))
        lvwCodeList.ColumnHeaders(i).Text = Trim(Replace(lvwCodeList.ColumnHeaders(i).Text, "▼", ""))
    Next
    
    With lvwCodeList
        .SortKey = ColumnHeader.index - 1
        .SortOrder = IIf(blnToggle(ColumnHeader.index - 1), lvwDescending, lvwAscending)
        .Sorted = True
        
        fraDLP.Caption = "검색 : " & ColumnHeader.Text & " 에서 찾기 "
        ColumnHeader.Text = ColumnHeader.Text & " " & IIf(.SortOrder = lvwAscending, "▲", "▼")
        
        blnToggle(ColumnHeader.index - 1) = IIf(blnToggle(ColumnHeader.index - 1), False, True)
    End With
    
    If lvwCodeList.ListItems.Count <> 0 Then
        lvwCodeList.ListItems(1).Selected = True
        lvwCodeList.ListItems(1).EnsureVisible
    End If
End Sub

Private Sub lvwCodeList_DblClick()
'헤더 갯수만큼 리턴해준다.

    Dim strSelectedItem As String
    Dim strColumnHeaderWidth As String
    Dim ItemX As ListItem
    Dim i As Long
    
    If lvwCodeList.ListItems.Count < 1 Then Exit Sub
    '
    Set ItemX = lvwCodeList.SelectedItem
    strSelectedItem = ItemX.Text
    
    For i = 1 To lvwCodeList.ColumnHeaders.Count - 1
        strSelectedItem = strSelectedItem & mvarDivision & ItemX.SubItems(i)
    Next
    
    For i = 1 To lvwCodeList.ColumnHeaders.Count
        strColumnHeaderWidth = strColumnHeaderWidth & lvwCodeList.ColumnHeaders(i).Width & mvarDivision
    Next
    
    strColumnHeaderWidth = Mid(strColumnHeaderWidth, 1, Len(strColumnHeaderWidth) - 1)
    
    mvarReturnColumnHeaderWidth = strColumnHeaderWidth
    mvarReturnFormWidth = frmDLP002.Width
    mvarReturnFormHeight = frmDLP002.Height
    mvarReturnFormTop = frmDLP002.Top
    mvarReturnFormLeft = frmDLP002.Left
    
    RaiseEvent SelectedItem(strSelectedItem)
End Sub

Private Sub optCode_Click(index As Integer)
    txtSearch.SetFocus
End Sub

Private Sub lvwCodeList_KeyDown(KeyCode As Integer, Shift As Integer)
    If lvwCodeList.ListItems.Count < 1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Then
        Call lvwCodeList_DblClick
    End If
End Sub

Private Sub txtSearch_Change()
'찾기
    Dim strFindItem As String
    Dim itmFound As ListItem   ' FoundItem 변수입니다.
    Dim itmX As ListItem
    Dim i As Long
    Dim lngColNo As Long
        
    strFindItem = Trim(txtSearch.Text)
    
    '▲▼
    With lvwCodeList
        For i = 1 To .ColumnHeaders.Count
            If InStr(.ColumnHeaders(i).Text, "▲") > 0 Or InStr(.ColumnHeaders(i).Text, "▼") > 0 Then
                lngColNo = i
                Exit For
            End If
        Next
        
        If lngColNo = 1 Then
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
                If UCase((itmX.SubItems(lngColNo - 1)) Like UCase((strFindItem & "*"))) Then
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

Private Sub txtSearch_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtSearch.Text) <> "" Then
            Call lvwCodeList_DblClick
        End If
    End If
    
    If KeyCode = vbKeyUp Then
        If lvwCodeList.SelectedItem.index - 1 = 0 Then Exit Sub
        
        lvwCodeList.ListItems(lvwCodeList.SelectedItem.index - 1).Selected = True
        lvwCodeList.ListItems(lvwCodeList.SelectedItem.index).EnsureVisible
    End If
    
    If KeyCode = vbKeyDown Then
        If lvwCodeList.SelectedItem.index + 1 > lvwCodeList.ListItems.Count Then Exit Sub
        
        lvwCodeList.ListItems(lvwCodeList.SelectedItem.index + 1).Selected = True
        lvwCodeList.ListItems(lvwCodeList.SelectedItem.index).EnsureVisible
    End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyUp Then
        lvwCodeList.SelectedItem.index = lvwCodeList.SelectedItem.index - 1
    End If
    
    If KeyAscii = vbKeyDown Then
        lvwCodeList.SelectedItem.index = lvwCodeList.SelectedItem.index + 1
    End If
End Sub
