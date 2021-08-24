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
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
        End If
    Else
        giPartCnt = CInt(sBuf)
    End If

'<------ Part 셋팅
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
                MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
            End If
        End If
                
        gPartTable(i).sPartName = GetKeyValue(HKEY_CURRENT_USER, _
            "Software\SemiLIS\Program Config\Part.Setting\Part." & CStr(i), "PartNm")
        
        If Trim$(gPartTable(i).sPartName) = "" Then
            If i = 1 Then
                bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Part.Setting\Part." & CStr(i), "PartNm", "생화학")
                gPartTable(i).sPartName = "생화학"
            ElseIf i = 2 Then
                bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Part.Setting\Part." & CStr(i), "PartNm", "혈액학")
                gPartTable(i).sPartName = "혈액학"
            ElseIf i = 3 Then
                bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Part.Setting\Part." & CStr(i), "PartNm", "혈청학")
                gPartTable(i).sPartName = "혈청학"
            ElseIf i = 4 Then
                bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Part.Setting\Part." & CStr(i), "PartNm", "뇨화학")
                gPartTable(i).sPartName = "뇨화학"
            Else
                bRetVal = UpdateKey(HKEY_CURRENT_USER, _
                    "Software\SemiLIS\Program Config\Part.Setting\Part." & CStr(i), "PartNm", "미정학부")
                gPartTable(i).sPartName = "미정학부"
            End If
            
            If bRetVal = True Then
            Else
                MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
            End If
        End If
    Next
End Sub


