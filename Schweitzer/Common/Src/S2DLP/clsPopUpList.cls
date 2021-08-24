VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPopUpList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'기존의 PopupList의 기능 추가판
'2003/10/21 Coding By legends

Public Event SelectedItem(ByVal pSelectedItem As String)

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private WithEvents objPop  As frmDLP002
Attribute objPop.VB_VarHelpID = -1

Private mvarFormCaption As String
Private mvarFormWidth As Long
Private mvarFormHeight As Long
Private mvarFormTop As Long
Private mvarFormLeft As Long

Private mvarColumnHeaderText As String    '헤더명
Private mvarColumnHeaderWidth As String       '컬럼크기
Private mvarColumnHeaderAlign As String
Private mvarHideColumnHeaders As Boolean     '컬럼 헤더 숨기기

Private mvarFontName As String
Private mvarFontSize As String

Private mvarReturnColumnHeaderWidth As String
Private mvarReturnFormWidth As Long
Private mvarReturnFormHeight As Long
Private mvarReturnFormTop As Long
Private mvarReturnFormLeft As Long

Private mvarSqlStmt As String       '쿼리
Private mvarRecordSet As Object     '레코드셋 개체
Private mvarDivision As String
Private mvarSelectedString As String
Private mvarSelectedItems As Variant     '배열로 넘길 놈

Private mvarSearchTool As Boolean   '검색 기능 여부(기본값은 검색기능 있게... True)
Private mvarClick As Boolean    '한번 클릭으로 할것인지 더블 클릭으로 선택할 것인지 체크(기본값은 더블클릭, False)

Public Sub SetRecordSet(ByRef pRecordSet As Object)
'메소드를 사용한 레코드 셋 넘겨받기위한 것
    Set mvarRecordSet = pRecordSet
End Sub

Public Property Let FormCaption(ByVal vData As String)
    mvarFormCaption = vData
End Property

Public Property Get FormCaption() As String
    FormCaption = mvarFormCaption
End Property

Public Property Let FormHeight(ByVal vData As Long)
    mvarFormHeight = vData
'    If mvarFormHeight < lngMeHeight Then mvarFormHeight = lngMeHeight
End Property

Public Property Get FormHeight() As Long
    FormHeight = mvarFormHeight
End Property

Public Property Let FormWidth(ByVal vData As Long)
    mvarFormWidth = vData
'    If mvarFormWidth < lngMeWidth Then mvarFormWidth = lngMeWidth
End Property

Public Property Get FormWidth() As Long
    FormWidth = mvarFormWidth
End Property

Public Property Let FormTop(ByVal vData As Long)
    mvarFormTop = vData
End Property

Public Property Get FormTop() As Long
    FormTop = mvarFormTop
End Property

Public Property Let FormLeft(ByVal vData As Long)
    mvarFormLeft = vData
End Property

Public Property Get FormLeft() As Long
    FormLeft = mvarFormLeft
End Property

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

Public Property Let HideColumnHeaders(ByVal vData As Boolean)
    mvarHideColumnHeaders = vData
End Property

Public Property Get HideColumnHeaders() As Boolean
    HideColumnHeaders = mvarHideColumnHeaders
End Property

Public Property Let FontName(ByVal vData As String)
    mvarFontName = vData
End Property

Public Property Get FontName() As String
    FontName = mvarFontName
End Property

Public Property Let FontSize(ByVal vData As String)
    mvarFontSize = vData
End Property

Public Property Get FontSize() As String
    FontSize = mvarFontSize
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

Public Property Let SelectedString(ByVal vData As String)
    mvarSelectedString = vData
End Property

Public Property Get SelectedString() As String
    SelectedString = mvarSelectedString
End Property

Public Property Let SelectedItems(ByVal vData As Variant)
    mvarSelectedItems = vData
End Property

Public Property Get SelectedItems() As Variant
    SelectedItems = mvarSelectedItems
End Property

Public Property Let SearchTool(ByVal vData As Boolean)
    mvarSearchTool = vData
End Property

Public Property Get SearchTool() As Boolean
    SearchTool = mvarSearchTool
End Property

Public Property Let Click(ByVal vData As Boolean)
    mvarClick = vData
End Property

Public Property Get Click() As Boolean
    Click = mvarClick
End Property

Public Sub LoadPopUp(Optional ByVal pSqlStmt As String = "")
    Dim p As POINTAPI
    Dim FormTop As Long
    Dim FormLeft As Long
    
    Call GetCursorPos(p)
    FormTop = IIf(FormTop = 0, p.Y * Screen.TwipsPerPixelY, FormTop)
    FormLeft = IIf(FormLeft = 0, p.X * Screen.TwipsPerPixelX, FormLeft)
    
    FormTop = IIf(mvarFormTop = 0, FormTop, mvarFormTop)
    FormLeft = IIf(mvarFormLeft = 0, FormLeft, mvarFormLeft)
    
    Set objPop = frmDLP002
    
    With objPop
        .ColumnHeaderText = mvarColumnHeaderText
        .ColumnHeaderWidth = mvarColumnHeaderWidth
        .ColumnHeaderAlign = mvarColumnHeaderAlign
        
        
        .lvwCodeList.Font.Name = mvarFontName
        .lvwCodeList.Font.Size = mvarFontSize
        .lvwCodeList.HideColumnHeaders = mvarHideColumnHeaders
        
        .SqlStmt = IIf(pSqlStmt = "", mvarSqlStmt, pSqlStmt)
        Set .RecordSet = mvarRecordSet
        .Caption = mvarFormCaption
        .Top = FormTop: .Left = FormLeft
        .Division = mvarDivision
        .fraDLP.Visible = IIf(mvarSearchTool, True, False)
        .Click = mvarClick
        
'        If mvarFormWidth > lngMeWidth Then
        If mvarFormWidth <> 0 Then
            .Width = mvarFormWidth
            .lvwCodeList.Width = .ScaleWidth
            .fraDLP.Width = .lvwCodeList.Width
            .txtSearch.Width = .fraDLP.Width - 150
        End If
        '
'        If mvarFormHeight > lngMeHeight Then
        If mvarFormHeight <> 0 Then
            .Height = mvarFormHeight
            .lvwCodeList.Height = IIf(mvarSearchTool, .ScaleHeight - (.fraDLP.Height + 75), .ScaleHeight)
            .fraDLP.Top = .lvwCodeList.Top + .lvwCodeList.Height + 75
        End If
        
        '위치조정--- 혹시 화면 밖을 나갈까봐서...
        If .Left < 0 Then .Left = .Width
        If .Top < 0 Then .Top = .Height
        If (.Left + .Width) > Screen.Width Then .Left = .Left - .Width
        If (.Top + .Height) > Screen.Height Then .Top = .Top - .Height
        
        .Show vbModal
    End With
End Sub

Private Sub Class_Initialize()
    mvarFormCaption = "코드 찾기"
    mvarFormWidth = 3615
    mvarFormHeight = 4125
    mvarFormTop = 0
    mvarFormLeft = 0
    
    mvarColumnHeaderText = ""   '헤더명
    mvarColumnHeaderWidth = ""      '컬럼크기
    mvarColumnHeaderAlign = ""
    mvarHideColumnHeaders = False
    
    mvarFontName = "굴림"
    mvarFontSize = "9"
    
    mvarReturnColumnHeaderWidth = ""
    mvarReturnFormWidth = 0
    mvarReturnFormHeight = 0
    mvarReturnFormTop = 0
    mvarReturnFormLeft = 0
    
    mvarSqlStmt = ""      '쿼리
    Set mvarRecordSet = Nothing    '레코드셋 개체
    mvarDivision = ""
    mvarSelectedString = ""
    mvarSelectedItems = Split(String(100, ","), ",")    '선택한 문자열의 기본값 설정을 위해서..
    
    mvarSearchTool = True
    mvarClick = False
End Sub

Private Sub objPop_SelectedItem(ByVal pSelectedItem As String)
    
    mvarReturnColumnHeaderWidth = objPop.ReturnColumnHeaderWidth
    mvarReturnFormWidth = objPop.ReturnFormWidth
    mvarReturnFormHeight = objPop.ReturnFormHeight
    mvarReturnFormTop = objPop.ReturnFormTop
    mvarReturnFormLeft = objPop.ReturnFormLeft
    
    mvarSelectedString = pSelectedItem
    mvarSelectedItems = Split(pSelectedItem, mvarDivision)
    
    Unload objPop
    Set objPop = Nothing
    RaiseEvent SelectedItem(pSelectedItem)
End Sub
