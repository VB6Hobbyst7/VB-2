Attribute VB_Name = "modKBSMC"
'
'   강북삼성병원
'
Option Explicit

'일반오더조회
Declare Function ExaminfoList2 Lib "C:\UniHis\DLL\P_SLDLL.dll" _
                            (ByVal sBarCd As String, ByRef vList As Variant) As Integer
'신검용 오더조회...2005/11/23
Declare Function ExaminfoList22 Lib "C:\UniHis\DLL\P_SLDLL.dll" _
                            (ByVal sBarCd As String, ByRef vList As Variant) As Integer
'LH-750 오더조회
Declare Function ExaminfoList3 Lib "C:\UniHis\DLL\P_SLDLL.dll" _
                            (ByVal sDate1 As String, ByVal sDate2 As String, ByRef vList As Variant) As Integer
'작업일자/WORKNO별 조회...2005/1/26 Add
Declare Function ExaminfoList4 Lib "C:\UniHis\DLL\P_SLDLL.dll" _
                            (ByVal sWDate As String, ByVal sWorkNo As String, ByRef vList As Variant) As Integer
'결과등록
Declare Function ResultList2 Lib "C:\UniHis\DLL\P_SLDLL.dll" _
                            (ByVal sEqCd As String, ByVal iCnt As Integer, ByVal sRstData As String) As Integer
'신검 결과등록
Declare Function ResultList22 Lib "C:\UniHis\DLL\P_SLDLL.dll" _
                            (ByVal sEqCd As String, ByVal iCnt As Integer, ByVal sRstData As String) As Integer
'서버 FLAG UPDATE
Declare Function FlagUpdate Lib "C:\UniHis\DLL\P_SLDLL.dll" (ByVal sBarCd As String) As Integer

'Init & Close
Declare Function TuxedoInit Lib "C:\UniHis\DLL\P_SLDLL.dll" (ByVal sUserNm$, ByVal sPara$) As Integer
Declare Function TuxedoTerm Lib "C:\UniHis\DLL\P_SLDLL.dll" () As Integer


Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_CLOSE = &H10


Public Function GetFRS01HWnd() As Long
    On Error GoTo ErrRtn
   
    Dim sBuf$
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "FRS01.HWnd")
    If Trim(sBuf) <> "" Then
        GetFRS01HWnd = CLng(sBuf)
    End If
    
ErrRtn:
    If Err <> 0 Then
        ViewMsg "GetFRS01HWnd - Err(" & Err.Description & ")"
    End If
End Function


Public Sub SetFRS01HWnd(ByVal lHWnd As Long)
    On Error GoTo ErrRtn
    
    Dim bRet    As Boolean
    
    bRet = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "FRS01.HWnd", Trim(lHWnd))
    
    If bRet <> True Then
        MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!", vbInformation
    End If
        
ErrRtn:
    If Err <> 0 Then
        MsgBox "SetFRS01HWnd - Err(" & Err.Description & ")", vbExclamation
    End If
End Sub


