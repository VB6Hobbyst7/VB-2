Attribute VB_Name = "ModFunction"
Public Sub Sub_MsgBox(vstring As String, vKind As Integer)
    
    '�޼��� �ڽ�
    Select Case vKind
        
        Case 1  '(!����) �����޼���
            MsgBox vstring, vbInformation, cMyCompany & "����͸�"
            funMsgBox = True
        Case 2  '(!�ﰢ) ���޼���
            MsgBox vstring, vbExclamation, cMyCompany & "����͸�"
            funMsgBox = True
        Case 3  '(X) �����޼���
            MsgBox vstring, vbCritical, cMyCompany & "����͸�"
            funMsgBox = True
        Case 4  'Ȯ��, ���
            If MsgBox(vstring, vbOKCancel, cMyCompany & "����͸�") = vbOK Then
                funMsgBox = True
            Else
                funMsgBox = False
            End If
        
    End Select

End Sub
