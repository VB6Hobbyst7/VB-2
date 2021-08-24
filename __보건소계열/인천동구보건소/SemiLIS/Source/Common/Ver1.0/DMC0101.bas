Attribute VB_Name = "DMC0101"
Option Explicit

Public Function fGetCurDSN() As String
    Dim sBuf$
    Dim bRetVal As Boolean
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\DSN", "")
    
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\DSN", "", "SemiLIS")
        
        If bRetVal = True Then
            fGetCurDSN = "SemiLIS"
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
            fGetCurDSN = "SemiLIS"
        End If
    Else
        fGetCurDSN = sBuf
    End If
End Function

Public Function fGetCurTestItemNmCfg() As String
    Dim sBuf As String
    Dim bRetVal As Boolean
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\Cur.Cfg", "TestItemNm Config")
    
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\Cur.Cfg", "TestItemNm Config", "T")
        'T : TestItemNm
        'P : PrintNm
        
        If bRetVal = True Then
            fGetCurTestItemNmCfg = "T"
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
            fGetCurTestItemNmCfg = "T"
        End If
    Else
        fGetCurTestItemNmCfg = sBuf
    End If
End Function

Public Function fGetCurPrintFlagCfg() As String
    Dim sBuf As String
    Dim bRetVal As Boolean
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\Cur.Cfg", "PrintFlag Config")
    
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\Cur.Cfg", "PrintFlag Config", "|||")
        
        If bRetVal = True Then
            fGetCurPrintFlagCfg = "|||"
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
            fGetCurPrintFlagCfg = "|||"
        End If
    Else
        fGetCurPrintFlagCfg = sBuf
    End If
End Function

Public Function fGetCurPrintPriority() As String
    Dim sBuf As String
    Dim bRetVal As Boolean
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\Cur.Cfg", "PrintPriority")
    
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, "Software\SemiLIS\Program Config\Cur.Cfg", "PrintPriority", "R")
        
        If bRetVal = True Then
            fGetCurPrintPriority = "R"
        Else
            MsgBox "������Ʈ��Ű�� �ʱ�ȭ �۾��� ������ �߻��߽��ϴ�!!"
            fGetCurPrintPriority = "R"
        End If
    Else
        fGetCurPrintPriority = sBuf
    End If
End Function

