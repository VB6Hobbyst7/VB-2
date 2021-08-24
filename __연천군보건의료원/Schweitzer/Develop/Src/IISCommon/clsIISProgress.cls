VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISProgress.cls
'   내  용  : Statusbar에 진행상태를 표시하는 클래스
'   버  전  :
'-----------------------------------------------------------------------------'
Option Explicit

'## 상태바의 패널의 크기
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'## AppearanceEnum
Public Enum AppearanceEnum
    pbFlat
    pb3d
End Enum

'## BorderStyleEnum
Public Enum BorderStyleEnum
    pbNone
    pbFixedSingle
End Enum

'## HAlignEnum
Public Enum HAlignEnum
    pbLeft
    pbCenter
    pbRight
End Enum

Private mPicProgress    As VB.PictureBox        'PictureBox

Private mDesForm        As Form                 '대상폼
Private mStatusBar      As Object               'StatusBar
Private mAppearance     As AppearanceEnum       'Appearance
Private mBorderStyle    As BorderStyleEnum      'BorderStyle
Private mHAlign         As HAlignEnum           '수평정렬
Private mMessage        As String               '표시할 메시지
Private mForeColor      As Long                 'ForeColor
Private mBackColor      As Long                 'BackColor
Private mPanelPos       As Long                 'StatusBar에서 ProgressBar가 표시될 Panel 위치
Private mFontSize       As Long                 '폰트크기
Private mMin            As Long                 'Min
Private mMax            As Long                 'Max
Private mValue          As Long                 'Value

Private mIsCreated      As Boolean              'ProgressBar의 생성여부

Public Property Let DesForm(ByRef vData As Object)
    Set mDesForm = vData
End Property

Public Property Let StatusBar(ByRef vData As Object)
    Set mStatusBar = vData
End Property

Public Property Let Appearance(ByVal vData As AppearanceEnum)
    mAppearance = vData
End Property

Public Property Let BorderStyle(ByVal vData As BorderStyleEnum)
    mBorderStyle = vData
End Property

Public Property Let HAlign(ByVal vData As HAlignEnum)
    mHAlign = vData
End Property

Public Property Let Message(ByVal vData As String)
    mMessage = vData
End Property

Public Property Let ForeColor(ByVal vData As Long)
    ForeColor = vData
End Property

Public Property Let BackColor(ByVal vData As Long)
    BackColor = vData
End Property

Public Property Let PanelPos(ByVal vData As Long)
    mPanelPos = vData - 1
End Property

Public Property Let FontSize(ByVal vData As Long)
    mFontSize = vData
End Property

Public Property Let Min(ByVal vData As Long)
    mMin = vData
End Property

Public Property Let Max(ByVal vData As Long)
    mMax = vData
End Property

Public Property Get Value() As Long
    Value = mValue
End Property

Public Property Let Value(ByVal vData As Long)
    Dim lngValue As Long
    
    If mStatusBar Is Nothing Then
        '## NOTE:강제적으로 에러를 발생
        Exit Property
    End If
    
    mValue = vData
    lngValue = (mValue / (mMax - mMin + 1)) * 100
    If lngValue < 0 Then lngValue = 1
    If lngValue > 100 Then lngValue = 100
    Call SetProgress(lngValue)
End Property

'-----------------------------------------------------------------------------'
'   기능 : ProgressBar를 생성
'-----------------------------------------------------------------------------'
Public Function CreateProgress()
    Dim rt As RECT

    '## PictureBox의 기본속성 설정
    Set mDesForm = New frmIISSplash
    Set mPicProgress = mDesForm.Controls.Add("VB.PictureBox", "mPicProgress")
    With mPicProgress
        .AutoRedraw = True
        .FillStyle = 0
        .ForeColor = mForeColor
        .BackColor = mBackColor
        .BorderStyle = mBorderStyle
        .Appearance = mAppearance
        .FontSize = mFontSize
    End With

    '## PictureBox를 StatusBar에 붙힘
    SendMessage mStatusBar.hwnd, SB_GETRECT, mPanelPos, rt
    With rt
        .Left = (.Left * Screen.TwipsPerPixelX) + 10
        .Top = (.Top * Screen.TwipsPerPixelY) + 10
        .Right = (.Right * Screen.TwipsPerPixelX) - .Left - 20
        .Bottom = (.Bottom * Screen.TwipsPerPixelY) - .Top - 20
    End With
    
    With mPicProgress
        SetParent .hwnd, mStatusBar.hwnd
        .Move rt.Left, rt.Top, rt.Right, rt.Bottom
        .Visible = True
    End With
    mIsCreated = True
End Function

'-----------------------------------------------------------------------------'
'   기능 : ProgressBar의 상태를 Update
'   인수 :
'       1.pValue : %값
'-----------------------------------------------------------------------------'
Public Function SetProgress(ByVal pValue As Long)
    If mIsCreated = False Then Call CreateProgress
    
    mPicProgress.Cls
    mPicProgress.ScaleWidth = 100
    mPicProgress.DrawMode = 10
    Select Case mHAlign
        Case pbLeft
            mPicProgress.CurrentX = 0
        Case pbCenter
            mPicProgress.CurrentX = 50 - mPicProgress.TextWidth(mMessage) / 2
        Case pbRight
            mPicProgress.CurrentX = mPicProgress.ScaleWidth - mPicProgress.TextWidth(mMessage)
    End Select
    
    mPicProgress.CurrentY = ((mPicProgress.ScaleHeight - mPicProgress.TextHeight(mMessage))) / 2
    If mMessage = "" Then
        mPicProgress.Print Format$(pValue, "###") & "%"
    Else
        mPicProgress.Print mMessage
    End If
    mPicProgress.Line (0, 0)-(pValue, mPicProgress.ScaleHeight), , BF
    mPicProgress.Refresh
End Function

Private Sub Class_Initialize()
    '## PictureBox의 기본속성 설정
    mForeColor = &HFF0000
    mBackColor = &H8000000F
    mBorderStyle = pbNone
    mAppearance = pb3d
    mHAlign = pbCenter
    mFontSize = 9
    mPanelPos = 0
    mMin = 1
    mMax = 100
End Sub

Private Sub Class_Terminate()
    If Not mDesForm Is Nothing Then
        Unload mDesForm
        Set mDesForm = Nothing
    End If
    Set mPicProgress = Nothing
End Sub

