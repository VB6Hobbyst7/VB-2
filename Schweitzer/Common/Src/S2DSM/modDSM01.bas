Attribute VB_Name = "modDSM01"
Option Explicit


'�ӽ� Test��  (By M.G.Choi)'=
Public GblUser As String   '=
Public GblEdit As Boolean  '=
'============================

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFilename As String) As Long

Public Const INIT_USER_SEC = "USER"
Public Const INIT_UID_KEY = "UID"
Public Const INIT_UNM_KEY = "UNM"
Public Const INIT_PWD_KEY = "PWD"
Public Const gintMAX_SIZE = 255

Public Sub Dithering(vObj As PictureBox)
    Dim intLoop As Integer
    
    vObj.DrawStyle = vbInsideSolid
    vObj.DrawMode = vbCopyPen
    vObj.ScaleMode = vbPixels
    vObj.DrawWidth = 4
    vObj.ScaleWidth = 100
    vObj.ScaleHeight = 255
    '--------------------------------------------------
    ' �Ķ���(0, 0, 255)���� ����������(0, 0, 0)����
    ' ���������� ĥ�� ������. ���� �����θ� ĥ�Ѵٴ�
    ' ������ �ִ�. �� ����� �ٲ��...
    '--------------------------------------------------
    For intLoop = 0 To 255
       vObj.Line (0, intLoop)-(100, intLoop - 1), RGB(intLoop, intLoop, intLoop), B
    Next intLoop
End Sub
