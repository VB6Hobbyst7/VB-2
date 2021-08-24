Attribute VB_Name = "VbHangul"
Option Explicit

Public Function LenH(ByVal strString As String) As Long

   LenH = LenB(StrConv(strString, vbFromUnicode))
'    LenH = LenB(strString)                      'MS SQL��

End Function

Public Function LeftH(ByVal strString As String, ByVal lngLength As Long) As String

    LeftH = StrConv(LeftB(StrConv(strString, vbFromUnicode), lngLength), vbUnicode)
'     LeftH = LeftB$(strString, lngLength)       'MS SQL��
     
End Function

Public Function RightH(ByVal strString As String, ByVal lngLength As Long) As String

    RightH = StrConv(RightB(StrConv(strString, vbFromUnicode), lngLength), vbUnicode)
'     RightH = RightB$(strString, lngLength)     'MS SQL��
     
End Function

Public Function MidH(ByVal strString As String, ByVal lngStart As Long, Optional ByVal lngLength As Variant) As String

    If IsMissing(lngLength) Then
        MidH = StrConv(MidB(StrConv(strString, vbFromUnicode), lngStart), vbUnicode)
    Else
        MidH = StrConv(MidB(StrConv(strString, vbFromUnicode), lngStart, lngLength), vbUnicode)
    End If

End Function

Public Function LPadH(ByVal strString As String, ByVal lngLength As Long) As String

'    LPadH = RightH(Space(lngLength) & strString, lngLength)
     LPadH = RightB$(Space(lngLength) & strString, lngLength)
     
End Function

Public Function RPadH(ByVal strString As String, ByVal lngLength As Long) As String

 '   RPadH = LeftH(strString & Space(lngLength), lngLength)
    RPadH = LeftB$(strString & Space(lngLength), lngLength)
    
End Function

Public Sub Check_MaxLength()

    On Error Resume Next
    
    If Not (TypeOf Screen.ActiveControl Is TextBox) Then Exit Sub
    
    If Screen.ActiveControl.MaxLength <> 0 Then
        If LenH(Screen.ActiveControl.Text) > Screen.ActiveControl.MaxLength Then
            SendKeys "{BS}"
        End If
    End If
    
End Sub

Public Function TxtMaxLen_Check(ArgObj As Object) As Boolean
    TxtMaxLen_Check = False
    If ArgObj.MaxLength < LenH(ArgObj.Text) Then
        MsgBox "�Է°��� �ʰ� �Ͽ����ϴ�." & vbNewLine & vbNewLine & _
               "�ѱ� " & ArgObj.MaxLength / 2 & " ��/���� " & ArgObj.MaxLength & _
               " �̳��� �Է��ϼ���", vbInformation + vbOKOnly, "Ȯ��"
        ArgObj.SelStart = 0
        ArgObj.SelLength = Len(ArgObj.Text)
        ArgObj.SetFocus
    Else
        TxtMaxLen_Check = True
    End If
    
End Function


