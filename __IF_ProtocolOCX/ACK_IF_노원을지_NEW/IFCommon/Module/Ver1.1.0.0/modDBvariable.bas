Attribute VB_Name = "modDBvariable"
Option Explicit

Public Const MAXORDERFIELD = 10
Public Const MAXRESULTFIELD = 12

Type ORDFIELDCFG
    sComponent As String
    sUse As String
    sStorageType As String
    sPath As String
    sFUse(MAXORDERFIELD) As String
    sFOrd(MAXORDERFIELD) As String
    sFSize(MAXORDERFIELD) As String
End Type

Type RSTFIELDCFG
    sComponent As String
    sUse As String
    sStorageType As String
    sPath As String
    sFUse(MAXRESULTFIELD) As String
    sFOrd(MAXRESULTFIELD) As String
    sFSize(MAXRESULTFIELD) As String
End Type

Public gOrdCfg As ORDFIELDCFG
Public gRstCfg As RSTFIELDCFG

Public Function fGetCurDSN_Old(ByVal sBuf As String) As String
'    Dim bRetVal As Boolean
'
'    sBuf = GetKeyValue(HKEY_CURRENT_USER, "Software\Ack_if\Interface Config\" & sBuf, "DSN")
'
'    If sBuf = "" Then
'        bRetVal = UpdateKey(HKEY_CURRENT_USER, "Software\Ack_if\Program Config\" & sBuf, "DSN", "IFDSN")
'
'        If bRetVal = True Then
'            fGetCurDSN = "IFDSN"
'        Else
'            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
'            fGetCurDSN = "IFDSN"
'        End If
'    Else
'        fGetCurDSN = sBuf
'    End If
End Function
Public Function fGetCurDSN(ByVal sBuf As String) As String
   
    Dim sRetVal As String
    
    'MS Access
    fGetCurDSN = "Driver={Microsoft Access Driver (*.mdb)};Dbq=[LOCALDB]"

    sRetVal = GetKeyValue(HKEY_CURRENT_USER, "Software\Ack_if\Interface Config\" & sBuf, "DSN")

    If sRetVal = "" Then
        sRetVal = App.Path & "\" & sBuf
        Call UpdateKey(HKEY_CURRENT_USER, "Software\Ack_if\Interface Config\" & sBuf, "DSN", sRetVal)
    End If
    fGetCurDSN = Replace(fGetCurDSN, "[LOCALDB]", Trim(sRetVal))

End Function
Public Function fGetCurTestItemNmCfg() As String
    Dim sBuf As String
    Dim bRetVal As Boolean
    
    sBuf = GetKeyValue(HKEY_CURRENT_USER, "Software\Ack_if\Program Config\Cur.Cfg", "TestItemNm Config")
    
    If sBuf = "" Then
        bRetVal = UpdateKey(HKEY_CURRENT_USER, "Software\Ack_if\Program Config\Cur.Cfg", "TestItemNm Config", "T")
        'T : TestItemNm
        'P : PrintNm
        
        If bRetVal = True Then
            fGetCurTestItemNmCfg = "T"
        Else
            MsgBox "레지스트리키의 초기화 작업에 에러가 발생했습니다!!"
            fGetCurTestItemNmCfg = "T"
        End If
    Else
        fGetCurTestItemNmCfg = sBuf
    End If
End Function

Public Function GetByOne(ByVal tStr As String, sOriginal As String) As String
    Dim Pos%
    
    Pos = InStr(tStr, "|")
    
    If Pos = 0 Then
    Else
        GetByOne = Trim$(Mid$(tStr, 1, Pos - 1))
        sOriginal = Trim$(Mid$(sOriginal, Pos + 1, Len(sOriginal) - Pos))
    End If
End Function

Public Function GetByOneUserSymbol(ByVal tStr As String, sOriginal As String, ByVal sUserSymbol As String) As String
    Dim Pos%

    Pos = InStr(tStr, sUserSymbol)

    If Pos = 0 Then
    Else
        GetByOneUserSymbol = Trim$(Mid$(tStr, 1, Pos - 1))
        sOriginal = Trim$(Mid$(sOriginal, Pos + 1, Len(sOriginal) - Pos))
    End If
End Function

Public Sub GetOrdRstCfg(ByVal sMachineCd As String)
    Dim sBuf$
    Dim i%
    
'Order.Use
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & sMachineCd, "Order.Use")
        
    gOrdCfg.sUse = sBuf
        
'Order.Component
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & sMachineCd, "Order.Component")
        
    gOrdCfg.sComponent = sBuf
        
'Order.Storage.Type
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & sMachineCd, "Order.Storage.Type")
        
    gOrdCfg.sStorageType = sBuf
    
'Order.Storage.Path
    If gOrdCfg.sStorageType = "" Then
        gOrdCfg.sPath = ""
    ElseIf gOrdCfg.sStorageType = "File" Then
        sBuf = GetKeyValue(HKEY_CURRENT_USER, _
            "Software\Ack_if\Interface Config\" & sMachineCd, "Order.FILE.Path")
            
        gOrdCfg.sPath = sBuf
    ElseIf gOrdCfg.sStorageType = "Database" Then
        sBuf = GetKeyValue(HKEY_CURRENT_USER, _
            "Software\Ack_if\Interface Config\" & sMachineCd, "Order.DB.DSN")
            
        gOrdCfg.sPath = sBuf
    Else
    End If
    
'Result.Use
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & sMachineCd, "Result.Use")
        
    gRstCfg.sUse = sBuf
    
'Result.Component
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & sMachineCd, "Result.Component")
        
    gRstCfg.sComponent = sBuf
    
'Result.Storage.Type
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & sMachineCd, "Result.Storage.Type")
        
    gRstCfg.sStorageType = sBuf
    
'Result.Storage.Path
    If gRstCfg.sStorageType = "" Then
        gRstCfg.sPath = ""
    ElseIf gRstCfg.sStorageType = "File" Then
        sBuf = GetKeyValue(HKEY_CURRENT_USER, _
            "Software\Ack_if\Interface Config\" & sMachineCd, "Result.FILE.Path")
            
        gRstCfg.sPath = sBuf
    ElseIf gRstCfg.sStorageType = "Database" Then
        sBuf = GetKeyValue(HKEY_CURRENT_USER, _
            "Software\Ack_if\Interface Config\" & sMachineCd, "Result.DB.DSN")
            
        gRstCfg.sPath = sBuf
    Else
    End If

'Order.Field.Use
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & sMachineCd, "Order.Field.Use")
    
    For i = 1 To MAXORDERFIELD
        gOrdCfg.sFUse(i) = GetByOne(sBuf, sBuf)
    Next
    
'Order.Field.FOrder
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & sMachineCd, "Order.Field.FOrder")
    
    For i = 1 To MAXORDERFIELD
        gOrdCfg.sFOrd(i) = Val(GetByOne(sBuf, sBuf))
    Next

'Order.Field.Size
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & sMachineCd, "Order.Field.Size")
    
    For i = 1 To MAXORDERFIELD
        gOrdCfg.sFSize(i) = Val(GetByOne(sBuf, sBuf))
    Next

'Result.Field.Use
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & sMachineCd, "Result.Field.Use")
    
    For i = 1 To MAXRESULTFIELD
        gRstCfg.sFUse(i) = GetByOne(sBuf, sBuf)
    Next
    
'Result.Field.FOrder
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & sMachineCd, "Result.Field.FOrder")
    
    For i = 1 To MAXRESULTFIELD
        gRstCfg.sFOrd(i) = Val(GetByOne(sBuf, sBuf))
    Next

'Result.Field.Size
    sBuf = GetKeyValue(HKEY_CURRENT_USER, _
        "Software\Ack_if\Interface Config\" & sMachineCd, "Result.Field.Size")
    
    For i = 1 To MAXRESULTFIELD
        gRstCfg.sFSize(i) = Val(GetByOne(sBuf, sBuf))
    Next

End Sub
