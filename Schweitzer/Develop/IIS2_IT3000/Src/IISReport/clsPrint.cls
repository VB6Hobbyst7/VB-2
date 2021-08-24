VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISPrint.cls
'   작성자  : 이상대
'   내  용  : Print 클래스
'   작성일  : 2004-12-17
'   버  전  :
'-----------------------------------------------------------------------------'

Option Explicit

'## Orientation Enum
Public Enum OrientationEnum
    ccPortrait = 1      '세로출력
    ccLandscape = 2     '가로출력
End Enum

'## Width Align Enum
Public Enum WAlignEnum
    ccLeft = 1          '왼쪽정렬
    ccCenter = 2        '가운데정렬
    ccRight = 3         '오른쪽정렬
End Enum

'## Height Align Enum
Public Enum HAlignEnum
    ccTop = 1           '위쪽정렬
    ccCenter = 2        '가운데정렬
    ccBottom = 3        '아래정렬
End Enum

Private mFont       As String       '메인폰트
Private mFontSize   As Long         '메인폰트크기
Private mStartXPos  As Single       'Stat X Position
Private mCurYPos    As Single       'Current Y Position
Private mLastYPos   As Single       'Last Y Position
Private mScaleWidth As Single       '종이의 가로여백
Private mSpace      As Single       '라인사이의 여백
Private mPageCnt    As Long         'Page Count

Private mTitle      As String       '레포트 타이틀
Private mTFont      As String       '타이틀 폰트
Private mTFontSize  As Long         '타이틀 폰트 사이즈
Private mTWAlign    As WAlignEnum   '타이틀의 가로정렬

Private mHeader     As String       '서브 타이틀
Private mColumns    As String       '컬럼 타이틀
Private mContent    As String       '레포트 내용

'-----------------------------------------------------------------------------'
'   기능 : 출력
'   인수 :
'       - pMode     : 출력방향
'       - pFont     : 폰트
'       - pFontSize : 폰트크기
'       - pStartX   : Start X Position
'       - pStartY   : Start Y Position
'       - pSpace    : 줄간격
'-----------------------------------------------------------------------------'
Public Sub PrintDoc(ByVal pMode As OrientationEnum, ByVal pFont As String, _
                    ByVal pFontSize As Long, Optional ByVal pStartX As Single = 5, _
                    Optional ByVal pStartY As Single = 10, Optional ByVal pSpace As Single = 6)

    '## 1.프린터 초기화
    mFont = pFont: mFontSize = pFontSize
    mStartXPos = pStartX: mCurYPos = pStartY: mSpace = pSpace
    
    With Printer
        .Font = mFont
        .FontSize = mFontSize
        .Orientation = pMode
        .ScaleMode = vbMillimeters
        
        mScaleWidth = .ScaleWidth
        mLastYPos = .ScaleHeight
    End With
    
    '## 2.Title 인쇄
    Call PrintTitle
    
    '## 3.Sub Title 인쇄
    
    '## 4.Column Header 인쇄
    
    '## 5.Content 인쇄
    
    Printer.EndDoc
End Sub

'-----------------------------------------------------------------------------'
'   기능 : Title 인쇄
'-----------------------------------------------------------------------------'
Private Sub PrintTitle()
    '## 타이틀, 폰트, 가로정렬
End Sub

'-----------------------------------------------------------------------------'
'   기능 : Title 설정
'   인수 :
'       -
'-----------------------------------------------------------------------------'
Public Sub SetTitle(ByVal pTitle As String, Optional pFont As String = "", _
                    Optional pFontSize As Long = -1, _
                    Optional pWidthAlign As WAlignEnum = WAlignEnum.ccCenter)
    
    '## 타이틀 설정 - 기본 굴림체, 20Size, 가운데정렬
    mTitle = pTitle
    mTFont = IIf(pFont = "", "굴림체", pFont)
    mTFontSize = IIf(pFontSize = -1, 20, pFontSize)
    mTWAlign = pWidthAlign
End Sub

Private Sub RowPrint(ByVal pContent As String, ByVal pX As Single, ByVal pY As String, _
                     ByVal pWidth As Single, Optional ByVal pWAlign As WAlignEnum = WAlignEnum.ccLeft, _
                     Optional ByVal pHAlign As HAlignEnum = HAlignEnum.ccCenter, _
                     Optional ByVal pAddLine As Boolean = False)

    '## 가로정렬
    Select Case pWAlign
        Case WAlignEnum.ccLeft   '왼쪽정렬
            Printer.CurrentX = pX + 0.5
        Case WAlignEnum.ccCenter '가운데정렬
            Printer.CurrentX = pX + (pWidth - Printer.TextWidth(pContent)) / 2
        Case WAlignEnum.ccRight  '오른쪽정렬
            Printer.CurrentX = pX + (pWidth - Printer.TextWidth(pContent)) - 0.5
    End Select
    
    '## 세로정렬
    Select Case pHAlign
        Case HAlignEnum.ccTop
            
        Case HAlignEnum.ccCenter
        Case HAlignEnum.ccBottom
    End Select
End Sub
