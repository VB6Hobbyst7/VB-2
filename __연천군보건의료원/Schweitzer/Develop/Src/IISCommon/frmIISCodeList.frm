VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIISCodeList 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "▶ CodeList"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3465
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin MSComctlLib.ListView lvwCodeList 
      Height          =   3060
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   5398
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
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "코드명"
         Object.Width           =   4022
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   585
      Left            =   0
      TabIndex        =   4
      Top             =   2970
      Width           =   3465
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFBFF&
         Height          =   330
         Left            =   1665
         ScrollBars      =   2  '수직
         TabIndex        =   1
         Top             =   180
         Width           =   1725
      End
      Begin VB.OptionButton optKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "코드명"
         Height          =   255
         Index           =   1
         Left            =   780
         TabIndex        =   3
         Top             =   210
         Width           =   855
      End
      Begin VB.OptionButton optKey 
         BackColor       =   &H00DBE6E6&
         Caption         =   "코드"
         Height          =   210
         Index           =   0
         Left            =   75
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmIISCodeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIISCodeList.frm
'   내  용  : 코드, 코드명을 Listview컨트롤에 표시
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

'## Parameter Type Enum
Public Enum ParamTypeEnum
    ccSql           'SQL Query가 전달된 경우
    ccRecordset     'Recordset이 전달된 경우
    ccCollection    '컬렉션 클래스가 전달된 경우
End Enum

Public Event SelectedItem(ByVal pSelItem As String)

Private mSQL        As String               'SQL Query
Private mRs         As ADODB.Recordset      'Recordset
Private mCol        As Object               'Collection Class
Private mParamType  As ParamTypeEnum        '전달된 인수의 형태

Public Property Let SQL(ByVal vData As String)
    mSQL = vData
End Property

Public Property Let Rs(ByRef vData As ADODB.Recordset)
    Set mRs = vData
End Property

Public Property Let Col(ByVal vData As Object)
    Set mCol = vData
End Property

Public Property Let ParamType(ByVal vData As ParamTypeEnum)
    mParamType = vData
End Property

Private Sub Form_Activate()
    txtSearch.SetFocus
    DoEvents
    
    Select Case mParamType
        Case ParamTypeEnum.ccSql:        Call LoadCodeBySql
        Case ParamTypeEnum.ccRecordset:  Call LoadCodeByRs
        Case ParamTypeEnum.ccCollection: Call LoadCodeByCol
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mRs = Nothing
    Set mCol = Nothing
    Set frmIISCodeList = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub lvwCodeList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static intOrder As Integer
    
    With lvwCodeList
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(intOrder = 0, lvwAscending, lvwDescending)
        .Sorted = True
        intOrder = (intOrder + 1) Mod 2
    End With
End Sub

Private Sub lvwCodeList_DblClick()
    Dim strSelectItem As String       '선택한 문자열
    
    With lvwCodeList
        If .ListItems.Count = 0 Then Exit Sub
        
        strSelectItem = .SelectedItem.Text & Chr(19) & .SelectedItem.SubItems(1)
        RaiseEvent SelectedItem(strSelectItem)
    End With
End Sub

Private Sub lvwCodeList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call lvwCodeList_DblClick
    End If
End Sub

Private Sub txtSearch_Change()
    Dim itmX        As ListItem
    Dim strSearch   As String     '찾는 문자열
    Dim i           As Long

    If lvwCodeList.ListItems.Count = 0 Then Exit Sub
    
    strSearch = Trim(txtSearch.Text)
    With lvwCodeList
        If optKey(0).Value Then
            For i = 1 To .ListItems.Count
                Set itmX = .ListItems(i)
                If UCase(itmX.Text) Like UCase(strSearch & "*") Then
                    itmX.Selected = True
                    itmX.EnsureVisible
                    Exit For
                End If
            Next i
        Else
            For i = 1 To .ListItems.Count
                Set itmX = .ListItems(i)
                If UCase(itmX.SubItems(1)) Like UCase(strSearch & "*") Then
                    itmX.Selected = True
                    itmX.EnsureVisible
                    Exit For
                End If
            Next i
        End If
    End With
    Set itmX = Nothing
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call lvwCodeList_DblClick
    ElseIf KeyCode = vbKeyUp Then
        lvwCodeList.SetFocus
        SendKeys "{UP}"
    ElseIf KeyCode = vbKeyDown Then
        lvwCodeList.SetFocus
        SendKeys "{DOWN}"
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 전달된 SQL문의 실행하여 Rs(0), Rs(1)의 내용을 Listview에 표시
'-----------------------------------------------------------------------------'
Private Sub LoadCodeBySql()
    Dim Rs      As ADODB.Recordset
    Dim itmX    As ListItem
    
    If mSQL = "" Then Exit Sub
    
On Error GoTo Errors
    Set Rs = mDbCon.Execute(mSQL, , adCmdText)
    If Rs.BOF Or Rs.EOF Then GoTo EndLine
    
    With lvwCodeList
        .ListItems.Clear
        Do Until Rs.EOF
            Set itmX = .ListItems.Add(, , Rs.Fields(0).Value)
            itmX.SubItems(1) = Rs.Fields(1).Value & ""
            
            Rs.MoveNext
        Loop
        Set itmX = Nothing
        
        If .ListItems.Count > 14 Then
            .ColumnHeaders(2).Width = .ColumnHeaders(2).Width - 200
        End If
    End With
    
EndLine:
    Rs.Close
    Set Rs = Nothing
    Exit Sub
    
Errors:
    Set Rs = Nothing
    mError.SetLog App.EXEName, "frmIISCodeList", "LoadCodeBySql", Err.Description, Now
    MsgBox mError.Description, vbCritical, "오류"
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 전달된 Recordset중 Rs(0), Rs(1)의 내용을 Listview에 표시
'-----------------------------------------------------------------------------'
Private Sub LoadCodeByRs()
    Dim itmX As ListItem
    
    If mRs Is Nothing Then Exit Sub
    
On Error GoTo Errors
    If mRs.BOF Or mRs.EOF Then GoTo EndLine
    
    With lvwCodeList
        .ListItems.Clear
        Do Until mRs.EOF
            Set itmX = .ListItems.Add(, , mRs.Fields(0).Value)
            itmX.SubItems(1) = mRs.Fields(1).Value & ""
            
            mRs.MoveNext
        Loop
        Set itmX = Nothing
        
        If .ListItems.Count > 14 Then
            .ColumnHeaders(2).Width = .ColumnHeaders(2).Width - 200
        End If
    End With
    
EndLine:
    mRs.Close
    Set mRs = Nothing
    Exit Sub
    
Errors:
    Set mRs = Nothing
    mError.SetLog App.EXEName, "frmIISCodeList", "LoadCodeByRs", Err.Description, Now
    MsgBox mError.Description, vbCritical, "오류"
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 전달된 컬렉션 클래스중 .Cd, .CdNm의 내용을 Listview에 표시
'-----------------------------------------------------------------------------'
Private Sub LoadCodeByCol()
    Dim Item As Variant
    Dim itmX As ListItem
    
    If mCol Is Nothing Then Exit Sub
    
    With lvwCodeList
        .ListItems.Clear
        For Each Item In mCol
            Set itmX = .ListItems.Add(, , Item.Cd)
            itmX.SubItems(1) = Item.CdNm
        Next
        
        If .ListItems.Count > 14 Then
            .ColumnHeaders(2).Width = .ColumnHeaders(2).Width - 200
        End If
    End With
    Set itmX = Nothing
End Sub
