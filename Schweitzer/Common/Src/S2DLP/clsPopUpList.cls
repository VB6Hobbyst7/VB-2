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

'������ PopupList�� ��� �߰���
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

Private mvarColumnHeaderText As String    '�����
Private mvarColumnHeaderWidth As String       '�÷�ũ��
Private mvarColumnHeaderAlign As String
Private mvarHideColumnHeaders As Boolean     '�÷� ��� �����

Private mvarFontName As String
Private mvarFontSize As String

Private mvarReturnColumnHeaderWidth As String
Private mvarReturnFormWidth As Long
Private mvarReturnFormHeight As Long
Private mvarReturnFormTop As Long
Private mvarReturnFormLeft As Long

Private mvarSqlStmt As String       '����
Private mvarRecordSet As Object     '���ڵ�� ��ü
Private mvarDivision As String
Private mvarSelectedString As String
Private mvarSelectedItems As Variant     '�迭�� �ѱ� ��

Private mvarSearchTool As Boolean   '�˻� ��� ����(�⺻���� �˻���� �ְ�... True)
Private mvarClick As Boolean    '�ѹ� Ŭ������ �Ұ����� ���� Ŭ������ ������ ������ üũ(�⺻���� ����Ŭ��, False)

Public Sub SetRecordSet(ByRef pRecordSet As Object)
'�޼ҵ带 ����� ���ڵ� �� �Ѱܹޱ����� ��
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
'�б� ����
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
        
        '��ġ����--- Ȥ�� ȭ�� ���� ���������...
        If .Left < 0 Then .Left = .Width
        If .Top < 0 Then .Top = .Height
        If (.Left + .Width) > Screen.Width Then .Left = .Left - .Width
        If (.Top + .Height) > Screen.Height Then .Top = .Top - .Height
        
        .Show vbModal
    End With
End Sub

Private Sub Class_Initialize()
    mvarFormCaption = "�ڵ� ã��"
    mvarFormWidth = 3615
    mvarFormHeight = 4125
    mvarFormTop = 0
    mvarFormLeft = 0
    
    mvarColumnHeaderText = ""   '�����
    mvarColumnHeaderWidth = ""      '�÷�ũ��
    mvarColumnHeaderAlign = ""
    mvarHideColumnHeaders = False
    
    mvarFontName = "����"
    mvarFontSize = "9"
    
    mvarReturnColumnHeaderWidth = ""
    mvarReturnFormWidth = 0
    mvarReturnFormHeight = 0
    mvarReturnFormTop = 0
    mvarReturnFormLeft = 0
    
    mvarSqlStmt = ""      '����
    Set mvarRecordSet = Nothing    '���ڵ�� ��ü
    mvarDivision = ""
    mvarSelectedString = ""
    mvarSelectedItems = Split(String(100, ","), ",")    '������ ���ڿ��� �⺻�� ������ ���ؼ�..
    
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