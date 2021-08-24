VERSION 5.00
Begin VB.UserControl uWait 
   BackStyle       =   0  '투명
   BorderStyle     =   1  '단일 고정
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   375
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   360
   ScaleWidth      =   375
   ToolboxBitmap   =   "uWait.ctx":0000
   Begin VB.Image Image1 
      Height          =   360
      Left            =   60
      Picture         =   "uWait.ctx":0312
      Top             =   0
      Width           =   225
   End
End
Attribute VB_Name = "uWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum PORT_TYPES
    PORT_IS_IRT
    PORT_IS_CARD
    PORT_IS_DLL
    PORT_IS_TRAN
End Enum
'응답전문
Private m_sRspMsg As String

'Title
Property Let p_sTitle(a_sTitle As String)
    fWait.Caption = a_sTitle
End Property
Property Let p_sDescription(a_sDesc As String)
    fWait.lblDesc = a_sDesc
End Property
Property Get p_sRspMsg() As String
    p_sRspMsg = m_sRspMsg
End Property
'Progress Bar
Public Sub f_vSetRange(a_iMin As Integer, a_iMax As Integer, Optional a_iStep As Integer = 1, Optional a_iInterval As Integer = 0)
    fWait.prgProgress.Min = a_iMin
    fWait.prgProgress.Max = Switch(a_iMax < 1, 1, a_iMax > 1, a_iMax)
    fWait.m_iStep = a_iStep
    If a_iInterval = 0 Then
        a_iInterval = 1
    End If
End Sub
'Start
Public Sub f_vStart(a_tPortType As PORT_TYPES, a_sSndMsg As String)
    fWait.f_vSendData a_tPortType, a_sSndMsg
    fWait.Show vbModal
    m_sRspMsg = fWait.p_sRspMsg
    Unload fWait
End Sub

Private Sub f_vControlSizeFix()
    '고정된크기의 Control
    UserControl.Height = 370
    UserControl.Width = 370
End Sub

Private Sub UserControl_Resize()
    f_vControlSizeFix
End Sub

Private Sub UserControl_Show()
    f_vControlSizeFix
End Sub

