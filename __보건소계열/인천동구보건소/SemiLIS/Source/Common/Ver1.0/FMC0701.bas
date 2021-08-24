Attribute VB_Name = "FMC0701"
Option Explicit

Public Sub InitializePart()
    Dim sBuf$
    Dim i%
    Dim bRetVal As Boolean
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\SemiLIS\Program Config\Part.Setting\Part.Cnt", "")

    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Part.Setting\Part.Cnt", "", "4")
        
        If bRetVal = True Then
            giPartCnt = 4
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
        End If
    Else
        giPartCnt = CInt(sBuf)
    End If

'<------ Part ����
    ReDim gPartTable(giPartCnt) As PartTBL
    
    For i = 1 To giPartCnt
        gPartTable(i).sPartInit = GetKeyValue(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Program Config\Part.Setting\Part." & CStr(i), "Init")
                
        If Trim$(gPartTable(i).sPartInit) = "" Then
            If i = 1 Then
                bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Part.Setting\Part." & CStr(i), "Init", "C")
                gPartTable(i).sPartInit = "C"
            ElseIf i = 2 Then
                bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Part.Setting\Part." & CStr(i), "Init", "H")
                gPartTable(i).sPartInit = "H"
            ElseIf i = 3 Then
                bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Part.Setting\Part." & CStr(i), "Init", "S")
                gPartTable(i).sPartInit = "S"
            ElseIf i = 4 Then
                bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Part.Setting\Part." & CStr(i), "Init", "U")
                gPartTable(i).sPartInit = "U"
            Else
                bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Part.Setting\Part." & CStr(i), "Init", "X")
                gPartTable(i).sPartInit = "X"
            End If
            
            If bRetVal = True Then
            Else
                MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
            End If
        End If
                
        gPartTable(i).sPartName = GetKeyValue(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Program Config\Part.Setting\Part." & CStr(i), "PartNm")
        
        If Trim$(gPartTable(i).sPartName) = "" Then
            If i = 1 Then
                bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Part.Setting\Part." & CStr(i), "PartNm", "��ȭ��")
                gPartTable(i).sPartName = "��ȭ��"
            ElseIf i = 2 Then
                bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Part.Setting\Part." & CStr(i), "PartNm", "������")
                gPartTable(i).sPartName = "������"
            ElseIf i = 3 Then
                bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Part.Setting\Part." & CStr(i), "PartNm", "��û��")
                gPartTable(i).sPartName = "��û��"
            ElseIf i = 4 Then
                bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Part.Setting\Part." & CStr(i), "PartNm", "��ȭ��")
                gPartTable(i).sPartName = "��ȭ��"
            Else
                bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Part.Setting\Part." & CStr(i), "PartNm", "�����к�")
                gPartTable(i).sPartName = "�����к�"
            End If
            
            If bRetVal = True Then
            Else
                MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
            End If
        End If
    Next
End Sub


