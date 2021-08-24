Attribute VB_Name = "ModFunction"
Public Sub Sub_MsgBox(vstring As String, vKind As Integer)
    
    '메세지 박스
    Select Case vKind
        
        Case 1  '(!동글) 정보메세지
            MsgBox vstring, vbInformation, cMyCompany & "모니터링"
            funMsgBox = True
        Case 2  '(!삼각) 경고메세지
            MsgBox vstring, vbExclamation, cMyCompany & "모니터링"
            funMsgBox = True
        Case 3  '(X) 오류메세지
            MsgBox vstring, vbCritical, cMyCompany & "모니터링"
            funMsgBox = True
        Case 4  '확인, 취소
            If MsgBox(vstring, vbOKCancel, cMyCompany & "모니터링") = vbOK Then
                funMsgBox = True
            Else
                funMsgBox = False
            End If
        
    End Select

End Sub
