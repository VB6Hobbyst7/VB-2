VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISMenus.cls
'   작성자  : 이상대
'   내  용  : Menus 컬렉션 클래스
'   작성일  : 2003-12-18
'   버  전  :
'       1. 1.1.7: 이상대(2005-04-13)
'-----------------------------------------------------------------------------'
Option Explicit

Private mMenus As Collection            '메뉴 컬렉션

Public Property Get Count() As Long
    Count = mMenus.Count
End Property

Public Property Get Item(ByVal pKey As String) As clsIISMenu
Attribute Item.VB_UserMemId = 0
    Set Item = mMenus(pKey)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mMenus.[_NewEnum]
End Property

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션에 항목추가
'-----------------------------------------------------------------------------'
Public Function Add(ByVal pKey As String, ByVal pVisible As Boolean) As clsIISMenu
    Dim objMenu As clsIISMenu
    
    '## 컬렉션에 있는 항목을 모두 삭제후 새항목 추가
    Set objMenu = New clsIISMenu
    With objMenu
        .Key = pKey
        .Visible = pVisible
    End With
    
    mMenus.Add objMenu, pKey
    
    Set Add = objMenu
    Set objMenu = Nothing
End Function

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션에서 해당요소 삭제
'-----------------------------------------------------------------------------'
Public Sub Remove(ByVal pKey As String)
    mMenus.Remove pKey
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컬렉션의 모든 요소삭제
'-----------------------------------------------------------------------------'
Public Sub RemoveAll()
    Dim i   As Long
    
    For i = mMenus.Count To 1 Step -1
        mMenus.Remove i
    Next i
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 키의 존재유무를 반환 - 1.1.7: 이상대(2005-04-13)
'   인수 :
'       - pKey : 버퍼순서
'   반환 : True(존재) / Flase(비존재)
'-----------------------------------------------------------------------------'
Public Function Exist(ByVal pKey As String) As Boolean
    Dim objMenu As clsIISMenu
    
    For Each objMenu In mMenus
        With objMenu
            If pKey = .Key Then
                Exist = True
                GoTo EndLine
            End If
        End With
    Next
    Exist = False

EndLine:
    Set objMenu = Nothing
End Function

Private Sub Class_Initialize()
    Set mMenus = New Collection
End Sub

Private Sub Class_Terminate()
    Set mMenus = Nothing
End Sub
