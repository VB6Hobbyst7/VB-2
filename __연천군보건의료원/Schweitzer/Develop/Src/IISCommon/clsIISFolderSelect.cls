VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISFolderSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISFolderSelect.cls
'   내  용  : 폴더선택폼을 보여주는 클래스
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

Public Event SelectedFolder(ByVal pSelFolder As String)

Private WithEvents mFolder As frmIISFolderSelect   '폴더선택 폼
Attribute mFolder.VB_VarHelpID = -1

Private mPath       As String           '선택한 경로
Private mCaption    As String           '폴더선택폼의 Caption
Private mLeft       As Long             '폴더선택폼의 Left 좌표
Private mTop        As Long             '폴더선택폼의 Top 좌표

Public Property Let Path(ByVal vData As String)
    mPath = vData
End Property

Public Property Let Caption(ByVal vData As String)
    mCaption = vData
End Property

Public Property Let Left(ByVal vData As Long)
    mLeft = vData
End Property

Public Property Let Top(ByVal vData As Long)
    mTop = vData
End Property

'-----------------------------------------------------------------------------'
'   기능 : 폴더선택폼 표시
'   인수 :
'       - pLeft : 폴더선택폼의 Left 좌표
'       - pTop  : 폴더선택폼의 Top 좌표
'-----------------------------------------------------------------------------'
Public Sub ShowFolderSelect(Optional ByVal pLeft As Long = 0, Optional ByVal pTop As Long = 0)
    '## Left, Top좌표를 검사
    Call SetWinPos(pLeft, pTop)
    
    If Dir$(mPath, vbDirectory) = "" Then
        mPath = ""
    End If
    
    '## FolderSelect 속성설정, 표시
    Set mFolder = frmIISFolderSelect
    With mFolder
        .Left = mLeft: .Top = mTop
        .Caption = "▶ " & mCaption
        .Path = mPath
        .ZOrder 0
        .Show vbModal
    End With
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 폴더선택폼 표시
'   인수 :
'       - pCaption  : 폴더선택폼의 Caption
'       - pPath     : 현재선택된 Path
'       - pLeft     : 폴더선택폼의 Left 좌표
'       - pTop      : 폴더선택폼의 Top 좌표
'-----------------------------------------------------------------------------'
Public Sub ShowFolderSelectX(ByVal pCaption As String, ByVal pPath As String, _
                             Optional ByVal pLeft As Long = 0, _
                             Optional ByVal pTop As Long = 0)
    '## Left, Top좌표를 검사
    Call SetWinPos(pLeft, pTop)
    
    If Dir$(pPath, vbDirectory) = "" Then
        mPath = ""
    End If
    
    '## FolderSelect 속성설정, 표시
    Set mFolder = frmIISFolderSelect
    With mFolder
        .Left = mLeft: .Top = mTop
        .Caption = "▶ " & pCaption
        .Path = pPath
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
'   기능 : 폴더선택폼에서 폴더를 선택했을때 발생하는 이벤트
'   인수
'       1.pSelFolder : 폴더선택폼에서 선택한 폴더의 전체경로
'-----------------------------------------------------------------------------'
Private Sub mFolder_SelectedFolder(ByVal pSelFolder As String)
    RaiseEvent SelectedFolder(pSelFolder)
    
    Unload mFolder
    Set mFolder = Nothing
End Sub

Private Sub Class_Terminate()
    Set mFolder = Nothing
End Sub
