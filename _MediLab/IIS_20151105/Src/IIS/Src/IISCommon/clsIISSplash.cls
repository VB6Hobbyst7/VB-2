VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIISSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------'
'   파일명  : clsIISSplash.cls
'   작성자  : 오세원
'   내  용  : Splash폼을 로드, 메시지 설정 클래스
'   작성일  : 2015-10-30
'   버  전  : 1.0.0
'-----------------------------------------------------------------------------'
Option Explicit

Private mSplashForm     As frmIISSplash     'Splash Form
Private mProjectNm      As String           '프로젝트명
Private mVersion        As String           '버전
Private mRegister       As String           '등록자
Private mMessage        As String           '메시지

Public Property Let ProjectNm(ByVal vData As String)
    mProjectNm = vData
'    mSplashForm.lblProjectNm = mProjectNm
End Property

Public Property Let Version(ByVal vData As String)
    mVersion = vData
    mSplashForm.lblVersion = "Version " & mVersion
End Property

Public Property Let Register(ByVal vData As String)
    mRegister = vData
    mSplashForm.lblRegister.Caption = "This version is registered to " & mRegister
End Property

Public Property Let Message(ByVal vData As String)
    mMessage = vData
    mSplashForm.lblMessage = mMessage & "..."
End Property

'-----------------------------------------------------------------------------'
'   기능 : Splash폼 로드
'   인수 : 표시할 메시지
'-----------------------------------------------------------------------------'
Public Sub LoadSplash()
    Dim objCom  As clsIISCommon
    
    '## 컴퓨터 이름 얻기
    Set objCom = New clsIISCommon
    If mRegister = "" Then
        mRegister = objCom.mGetComNm
        mSplashForm.lblRegister.Caption = "This version is registered to " & mRegister
    End If
    
    Set objCom = Nothing
    mSplashForm.Show
    mSplashForm.ZOrder 0
End Sub

'-----------------------------------------------------------------------------'
'   기능 : Splash폼에 메시지를 표시
'   인수 : 표시할 메시지
'-----------------------------------------------------------------------------'
Public Sub SetMsg(ByVal pMsg As String)
    mSplashForm.lblMessage = pMsg & "..."
End Sub

Private Sub Class_Initialize()
    Set mSplashForm = New frmIISSplash
End Sub

Private Sub Class_Terminate()
    Unload mSplashForm
    Set mSplashForm = Nothing
End Sub
