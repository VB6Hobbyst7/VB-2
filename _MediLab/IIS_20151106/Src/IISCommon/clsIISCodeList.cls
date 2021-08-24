VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISCodeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISCodeList.cls
'   작성자  : 오세원
'   내  용  : Code List 원도우를 보여주는 클래스
'   작성일  : 2015-10-30
'   버  전  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

Public Event SelectedItem(ByRef pSelItem As String)

Private WithEvents mCodeList As frmIISCodeList     'CodeList폼
Attribute mCodeList.VB_VarHelpID = -1

Private mLeft       As Long         'CodeList폼이 표시되는 Left좌표
Private mTop        As Long         'CodeList폼이 표시되는 Top좌표
Private mCaption    As String       'CodeList폼의 Caption
Private mHeaderCd   As String       'CodeList폼의 Listview의 코드에 해당하는 컬럼명
Private mHeaderCdNm As String       'CodeList폼의 Listview의 코드명에 해당하는 컬럼명

Public Property Get Left() As Long
    Left = mLeft
End Property

Public Property Let Left(ByVal vData As Long)
    mLeft = vData
End Property

Public Property Get Top() As Long
    Top = mTop
End Property

Public Property Let Top(ByVal vData As Long)
    mTop = vData
End Property

Public Property Get Caption() As String
    Caption = mCaption
End Property

Public Property Let Caption(ByVal vData As String)
    mCaption = "▶ " & vData
End Property

Public Property Get HeaderCd() As String
    HeaderCd = mHeaderCd
End Property

Public Property Let HeaderCd(ByVal vData As String)
    mHeaderCd = vData
End Property

Public Property Get HeaderCdNm() As String
    HeaderCdNm = mHeaderCdNm
End Property

Public Property Let HeaderCdNm(ByVal vData As String)
    mHeaderCdNm = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 전달된 SQL문의 실행하여 Rs(0), Rs(1)의 내용을 Listview에 표시
'   인수 :
'       1.pSQL  : 코드, 코드명을 가지는 SQL문
'       2.pLeft : CodeList폼이 표시되는 Left좌표
'       3.pTop  : CodeList폼이 표시되는 Top좌표
'-----------------------------------------------------------------------------'
Public Sub CodeListBySql(ByVal pSQL As String, Optional ByVal pLeft As Long = 0, _
                         Optional ByVal pTop As Long = 0)

    '## Left, Top좌표를 검사
    Call SetWinPos(pLeft, pTop)
    
    '## CodeList 속성설정, 표시
    Set mCodeList = frmIISCodeList
    With mCodeList
        .Left = mLeft: .Top = mTop
        .Caption = mCaption
        .lvwCodeList.ColumnHeaders(1).Text = mHeaderCd
        .lvwCodeList.ColumnHeaders(2).Text = mHeaderCdNm
        .ParamType = ccSql: .SQL = pSQL
        
        .ZOrder 0
        .Show vbModal
    End With
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 전달된 Recordset중 Rs(0), Rs(1)의 내용을 Listview에 표시
'   인수 :
'       1.pRs   : 코드, 코드명을 가지는 RecordSet
'       2.pLeft : CodeList폼이 표시되는 Left좌표
'       3.pTop  : CodeList폼이 표시되는 Top좌표
'-----------------------------------------------------------------------------'
Public Sub CodeListByRs(ByRef pRs As Recordset, Optional ByVal pLeft As Long = 0, _
                        Optional ByVal pTop As Long = 0)
    '## Left, Top좌표를 검사
    Call SetWinPos(pLeft, pTop)
    
    '## CodeList 속성설정, 표시
    Set mCodeList = frmIISCodeList
    With mCodeList
        .Left = mLeft: .Top = mTop
        .Caption = mCaption
        .lvwCodeList.ColumnHeaders(1).Text = mHeaderCd
        .lvwCodeList.ColumnHeaders(2).Text = mHeaderCdNm
        .ParamType = ccRecordset: .Rs = pRs
        
        .ZOrder 0
        .Show vbModal
    End With
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 전달된 컬렉션 클래스중 .Cd, .CdNm의 내용을 Listview에 표시
'   인수 :
'       1.pCol  : 컬렉션 클래스
'       2.pLeft : CodeList폼이 표시되는 Left좌표
'       3.pTop  : CodeList폼이 표시되는 Top좌표
'   필요사항 : 클래스는 반드는 Cd, CdNm Property를 구현해야 한다.
'-----------------------------------------------------------------------------'
Public Sub CodeListByCol(ByVal pCol As Object, Optional ByVal pLeft As Long = 0, _
                         Optional ByVal pTop As Long = 0)
    '## Left, Top좌표를 검사
    Call SetWinPos(pLeft, pTop)
    
    '## CodeList 속성설정, 표시
    Set mCodeList = frmIISCodeList
    With mCodeList
        .Left = mLeft: .Top = mTop
        .Caption = mCaption
        .lvwCodeList.ColumnHeaders(1).Text = mHeaderCd
        .lvwCodeList.ColumnHeaders(2).Text = mHeaderCdNm
        .ParamType = ccCollection: .Col = pCol
        
        .ZOrder 0
        .Show vbModal
    End With
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 폼의 위치를 전달하지 않은 경우 현재 마우스 포인터로 맞추어
'          CodeList폼의 Left, Top의 위치를 설정
'-----------------------------------------------------------------------------'
Private Sub SetWinPos(ByVal pLeft As Long, ByVal pTop As Long)
    Dim Point As POINTAPI
    
    Call GetCursorPos(Point)
    
    mLeft = IIf(pLeft = 0, Point.X * Screen.TwipsPerPixelX, pLeft)
    mTop = IIf(pTop = 0, Point.Y * Screen.TwipsPerPixelY, pTop)
End Sub

'-----------------------------------------------------------------------------'
'   기능 : CodeList폼에서 아이템을 선택했을때 발생하는 이벤트
'   인수
'       1.pSelItem : CodeList폼에서 선택한 아이템
'-----------------------------------------------------------------------------'
Private Sub mCodeList_SelectedItem(ByVal pSelItem As String)
    RaiseEvent SelectedItem(pSelItem)
    
    '## 이벤트 구현쪽에서 pSelItem=""을 입력하면 현재 CodeList폼을 그대로 유지
    If pSelItem <> "" Then
        Unload mCodeList
        Set mCodeList = Nothing
    End If
End Sub

Private Sub Class_Initialize()
    mHeaderCd = "코드"
    mHeaderCdNm = "코드명"
End Sub

Private Sub Class_Terminate()
    Set mCodeList = Nothing
End Sub

