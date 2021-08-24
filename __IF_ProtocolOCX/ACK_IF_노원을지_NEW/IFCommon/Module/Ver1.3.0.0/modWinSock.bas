Attribute VB_Name = "modWinSock"
Option Explicit

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_CLOSE = &H10
Public Function GetRegSvrHWnd() As Long
    On Error GoTo ErrRtn
   
    Dim sBuf$
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
                "Software\Ack_if\Interface Config\" & gsMachineCd, "AutoReg.HWnd")
    If Trim(sBuf) <> "" Then
        GetRegSvrHWnd = CLng(sBuf)
    End If
    
ErrRtn:
    If Err <> 0 Then
        ViewMsg "GetRegSvrHWnd - Err(" & Err.Description & ")"
    End If
End Function

Public Sub SetRegSvrHWnd(ByVal lHWnd As Long)
    On Error GoTo ErrRtn
    
    Dim bRet    As Boolean
    
    bRet = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\Ack_if\Interface Config\" & gsMachineCd, "AutoReg.HWnd", Trim(lHWnd))
    
    If bRet <> True Then
        MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!", vbInformation
    End If
        
ErrRtn:
    If Err <> 0 Then
        MsgBox "SetRegSvrHWnd - Err(" & Err.Description & ")", vbExclamation
    End If
End Sub

Public Sub SendResultSocket(ByVal iMode As Integer, ByVal sCRow As String, ByVal iRstCnt As Integer, _
                            ByVal sIFRstCd As String, ByVal sRst1 As String, ByVal sRst2 As String, _
                            Optional ByVal iCnt As Integer)
    On Error GoTo ErrHandler
    
    Dim vIFItemCnt, vTmp, vChk
    Dim i%, j%, k%, iExist%
    Dim sTmp$, sCIFRstCd$, sCRst1$, sCRst2$, sCFlag$
    Dim sWDate$, sWSeq$, sJDate$, sJGbn$, sJNo$, sRack$, sPos$, sRegNo$, sName$, sSex$, sEmer$, sReRun$, sOther$
    Dim sTIFSeq$, sTRst1$, sTRst2$, sTFlag$
    Dim sIFSeq$, sRtnVal$
    
    With gfIFDisplayForm.spdIntList
        sWDate = Format(frmInterface.dtpLabDate.Value, "YYYYMMDD")

        Call .GetText(1, CInt(sCRow), vTmp)
        sWSeq = CStr(vTmp)

        Call .GetText(3, CInt(sCRow), vTmp)
        sJDate = CStr(vTmp)
        
        Call .GetText(4, CInt(sCRow), vTmp)
        sJGbn = CStr(vTmp)
        
        Call .GetText(5, CInt(sCRow), vTmp)
        sJNo = CStr(vTmp)
        
        Call .GetText(6, CInt(sCRow), vTmp)
        sRack = CStr(vTmp)
        
        Call .GetText(7, CInt(sCRow), vTmp)
        sPos = CStr(vTmp)
        
        Call .GetText(8, CInt(sCRow), vTmp)
        sRegNo = CStr(vTmp)
        
        Call .GetText(9, CInt(sCRow), vTmp)
        sName = CStr(vTmp)
        
        Call .GetText(10, CInt(sCRow), vTmp)
        sSex = CStr(vTmp)
        
        Call .GetText(11, CInt(sCRow), vTmp)
        sEmer = CStr(vTmp)
        
        Call .GetText(12, CInt(sCRow), vTmp)
        sReRun = CStr(vTmp)
        
        Call .GetText(13, CInt(sCRow), vTmp)
        sOther = CStr(vTmp)
        

    'iMode = 1 ---> 한 샘플씩 LOCAL 등록
        Call .GetText(16, CInt(sCRow), vIFItemCnt)
        
        For i = 1 To CInt(vIFItemCnt)
            Call .GetText(16 + i, CInt(sCRow), vTmp)
            
            sTmp = CStr(vTmp)
            
            sIFSeq = GetByOne(sTmp, sTmp)  '검사항목코드
            sRst1 = GetByOne(sTmp, sTmp)
            sRst2 = GetByOne(sTmp, sTmp)
            
            sTIFSeq = sTIFSeq & sIFSeq & "|"
            sTRst1 = sTRst1 & sRst1 & "|"
            sTRst2 = sTRst2 & sRst2 & "|"
        Next i
    End With
    
    If Len(sJNo) < 11 Then
    Else
        '--- 결과등록 프로그램에 메세지 전송(2003/3/17 yk)
        Dim sSendMsg    As String
        
        sSendMsg = "R" & Chr(3) & sWDate & Chr(3) & sWSeq & Chr(3) & sJNo & Chr(3) _
                & sTIFSeq & Chr(3) & sTRst1 & Chr(4)
        
        With frmInterface.Winsock1
            If .State = sckConnected Then
                .SendData sSendMsg
            End If
        End With
        '-------------------------------------------------
    End If
    
    Exit Sub
ErrHandler:
    ViewMsg "SendResultSocket 오류 - (" & Err.Description & ")"
End Sub



