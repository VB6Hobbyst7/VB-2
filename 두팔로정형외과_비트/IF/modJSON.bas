Attribute VB_Name = "modJSON"
Option Explicit

Public mJsonData As String

Public Function getJsonVar(ByRef v_strData As String) As clsJSON

    Dim objResult As New clsJSON
    Dim objCurrent As clsJSON
    Dim i As Integer
    Dim IngStartPos As Long
    Dim IngEndPos As Long
    Dim IngLength As Long
    Dim strkey As String
    Dim strValue As String
    Dim strType As String
    Dim bMode As Boolean
    '-- 숫자값 처리용
    Dim bISNMode As Boolean
    Dim IngPrevStartPos As Long
    Dim IngPrevEndPos As Long
    
    mJsonData = ""
    bMode = False
    bISNMode = False
    IngStartPos = 0
    IngLength = Len(v_strData)
    
    Set objCurrent = objResult

    Do
        IngStartPos = InStr(IngStartPos + 1, v_strData, Chr$(34))
        IngEndPos = InStr(IngStartPos + 1, v_strData, Chr$(34))
        If IngEndPos = 0 Or IngStartPos = 0 Then Exit Do
    
        If bMode Then
            strValue = Mid$(v_strData, IngStartPos + 1, IngEndPos - IngStartPos - 1)
        Else
            If bISNMode = True Then
                strValue = mGetP(Mid$(v_strData, IngPrevStartPos + 2), 1, ",")
                bISNMode = False
            End If
            If strValue <> "" Then
                mJsonData = mJsonData & strkey & "@" & strValue & vbCr
            Else
                If strkey <> "" Then
                    mJsonData = mJsonData & strkey & "@" & strValue & vbCr
                End If
            End If
            
            strkey = Mid$(v_strData, IngStartPos + 1, IngEndPos - IngStartPos - 1)
            
        End If
        
        If strValue <> "" Then
            mJsonData = mJsonData & strkey & "@" & strValue & vbCr
        Else
            If strkey <> "" And strValue <> "" Then
                mJsonData = mJsonData & strkey & "@" & strValue & vbCr
            End If
        End If
        
        Select Case Mid$(v_strData, IngEndPos + 1, 1)
            Case ":"
                Select Case Mid$(v_strData, IngEndPos + 2, 1)
                    Case "{"
                        Set objCurrent = objCurrent.addChild(strkey)
                    Case Chr$(34)
                        bMode = True
                    Case "["
                        Set objCurrent = objCurrent.addChild(strkey).addChild()
                    Case ")"
                        'Stop
                    Case Else
                        If IsNumeric(Mid$(v_strData, IngEndPos + 2, 1)) Then
                            bISNMode = True
                        End If
                End Select
            Case ","
                Call objCurrent.addChild(strkey, strValue)
                strkey = ""
                strValue = ""
                bMode = False
    
            Case "}"
                Call objCurrent.addChild(strkey, strValue)
                Set objCurrent = objCurrent.getParent
                strkey = ""
                strValue = ""
                bMode = False
    
                If Mid$(v_strData, IngEndPos + 3, 1) = "{" Then
                    Set objCurrent = objCurrent.addChild()
                End If
        End Select
        IngStartPos = IngEndPos
        IngPrevStartPos = IngStartPos
        IngPrevEndPos = IngEndPos
    Loop
    
    Set getJsonVar = objResult

End Function

'''Public Function getJsonVarPT(ByRef v_strData As String) As clsJSON
'''
'''    Dim objResult As New clsJsonPT
'''    Dim objCurrent As clsJsonPT
'''    Dim i As Integer
'''    Dim IngStartPos As Long
'''    Dim IngEndPos As Long
'''    Dim IngLength As Long
'''    Dim strkey As String
'''    Dim strValue As String
'''    Dim strType As String
'''    Dim bMode As Boolean
'''    '-- 숫자값 처리용
'''    Dim bISNMode As Boolean
'''    Dim IngPrevStartPos As Long
'''    Dim IngPrevEndPos As Long
'''
'''    mJsonData = ""
'''    bMode = False
'''    bISNMode = False
'''    IngStartPos = 0
'''    IngLength = Len(v_strData)
'''
'''    Set objCurrent = objResult
'''
'''    Do
'''        IngStartPos = InStr(IngStartPos + 1, v_strData, Chr$(34))
'''        IngEndPos = InStr(IngStartPos + 1, v_strData, Chr$(34))
'''        If IngEndPos = 0 Or IngStartPos = 0 Then Exit Do
'''
'''        If bMode Then
'''            strValue = Mid$(v_strData, IngStartPos + 1, IngEndPos - IngStartPos - 1)
'''        Else
'''            If bISNMode = True Then
'''                strValue = mGetP(Mid$(v_strData, IngPrevStartPos + 2), 1, ",")
'''                bISNMode = False
'''            End If
'''            If strValue <> "" Then
'''                SetRawData Trim(strkey) & "@" & Trim(strValue)
'''                mJsonData = mJsonData & strkey & "@" & strValue & vbCr
'''            End If
'''            strkey = Mid$(v_strData, IngStartPos + 1, IngEndPos - IngStartPos - 1)
'''        End If
'''
'''        If strValue <> "" Then
'''            SetRawData Trim(strkey) & "@" & Trim(strValue)
'''            mJsonData = mJsonData & strkey & "@" & strValue & vbCr
'''        End If
'''
'''        Select Case Mid$(v_strData, IngEndPos + 1, 1)
'''            Case ":"
'''                Select Case Mid$(v_strData, IngEndPos + 2, 1)
'''                    Case "{"
'''                        Set objCurrent = objCurrent.addChild(strkey)
'''                    Case Chr$(34)
'''                        bMode = True
'''                    Case "["
'''                        Set objCurrent = objCurrent.addChild(strkey).addChild()
'''                    Case ")"
'''                        'Stop
'''                    Case Else
'''                        If IsNumeric(Mid$(v_strData, IngEndPos + 2, 1)) Then
'''                            bISNMode = True
'''                        End If
'''                End Select
'''            Case ","
'''                Call objCurrent.addChild(strkey, strValue)
'''                strkey = ""
'''                strValue = ""
'''                bMode = False
'''
'''            Case "}"
'''                Call objCurrent.addChild(strkey, strValue)
'''                Set objCurrent = objCurrent.getParent
'''                strkey = ""
'''                strValue = ""
'''                bMode = False
'''
'''                If Mid$(v_strData, IngEndPos + 3, 1) = "{" Then
'''                    Set objCurrent = objCurrent.addChild()
'''                End If
'''        End Select
'''        IngStartPos = IngEndPos
'''        IngPrevStartPos = IngStartPos
'''        IngPrevEndPos = IngEndPos
'''    Loop
'''
'''    Set getJsonVarPT = objResult
'''
'''End Function

Public Function JsonSend(strAction As String, strParam() As Variant) As Variant
    Dim strURL      As String
    Dim strHeader   As String
    Dim varPara()   As Variant
    Dim varVal()    As Variant
    Dim strVHDV     As String
    
    Select Case strAction
        Case "LOGIN"
        '워크조회
        Case "WORKLIST"
            strURL = gURL.WORKLIST
            strHeader = "srchMap"
    
            ReDim Preserve varPara(5) As Variant
            varPara(0) = "rcpnDt1"
            varPara(1) = "rcpnDt2"
            varPara(2) = "slipCd"
            varPara(3) = "workNo1"
            varPara(4) = "workNo2"
            varPara(5) = "exmnCd"
    
            ReDim Preserve varVal(5) As Variant
            varVal(0) = strParam(0)
            varVal(1) = strParam(1)
            varVal(2) = strParam(2)
            varVal(3) = strParam(3)
            varVal(4) = strParam(4)
            varVal(5) = strParam(5)
            
        '바코드조회
        Case "PATLIST"
            strURL = gURL.PATLIST
            strHeader = "srchMap"
    
            ReDim Preserve varPara(0) As Variant
            varPara(0) = "brcdLablNo"
    
            ReDim Preserve varVal(0) As Variant
            varVal(0) = strParam(0)
        
        '결과저장
        Case "PATSAVE"
            strURL = gURL.PATSAVE
            strHeader = "saveList"
            
            ReDim Preserve varPara(7) As Variant
            varPara(0) = "brcdLablNo"
            varPara(1) = "exmnCd"
            varPara(2) = "realRslt"
            varPara(3) = "viewRslt"
            varPara(4) = "eqpmCd"
            varPara(5) = "eqpmFlag"
            varPara(6) = "examDt"
            varPara(7) = "exmnId"
    
    '  brcdLablNo : "1820027311"
    '  exmnCd : "L3011"
    '  realRslt :"100"
    '  viewRslt : "100"
    '  eqpmCd : "011"
    '  eqpmFlag : "1"
    '  examDt : "20180504010101"
    '  exmnId : "test"
      
            ReDim Preserve varVal(7) As Variant
            varVal(0) = strParam(0)
            varVal(1) = strParam(1)
            varVal(2) = strParam(2)
            varVal(3) = strParam(3)
            varVal(4) = strParam(4)
            varVal(5) = strParam(5)
            varVal(6) = strParam(6)
            varVal(7) = strParam(7)
    End Select
    
    strURL = gHOSP.HOSPCD & strURL
    JsonSend = JSONRPC(strURL, strHeader, varPara, varVal, -1)
    
End Function

Public Function JSONRPC(URL$, JSONPostHeader$, P() As Variant, V() As Variant, Optional intAct As Integer) As String
    Dim http    As Object
    Dim i       As Integer
    Dim JSONPostBody$()
  
On Error GoTo RST
    
    Set http = CreateObject("Winhttp.WinHttpRequest.5.1")
      
    http.Open "POST", URL, False
    http.setRequestHeader "Content-Type", "application/json; charset=UTF-8"
    http.setRequestHeader "Accept", "application/json"
    ReDim Preserve JSONPostBody$(UBound(P))
    For i = 0 To UBound(P)
        JSONPostBody(i) = MakeJSONFromParams(P(i)) & ":" & MakeJSONFromParams(V(i))
    Next
    http.send "{" & Join(JSONPostBody, ",") & "}"
    JSONRPC = http.responseText
    Set http = Nothing

Exit Function
RST:
    JSONRPC = ""
    Set http = Nothing

End Function



Public Function MakeJSONFromParams(ByVal P) As String 'Helper-function for the above main-request-function
    Dim Tmp$
    
    Select Case VarType(P)
        Case vbString:        Tmp = """" & P & """"
        Case vbBoolean:       Tmp = IIf(P, "true", "false")
        Case vbEmpty, vbNull: Tmp = "null"
        Case Else:            Tmp = Str$(P)
    End Select
    
    MakeJSONFromParams = Tmp
    
End Function

Public Function MakeJSONArrayFromParams(ByVal PArr) As String 'Helper-function for the above main-request-function
    Dim Tmp$(), P
    
    Tmp = Split(vbNullString)
    
    For Each P In PArr
        ReDim Preserve Tmp(0 To UBound(Tmp) + 1)
        Select Case VarType(P)
            Case vbString:        Tmp(UBound(Tmp)) = """" & P & """"
            Case vbBoolean:       Tmp(UBound(Tmp)) = IIf(P, "true", "false")
            Case vbEmpty, vbNull: Tmp(UBound(Tmp)) = "null"
            Case Else:            Tmp(UBound(Tmp)) = Str$(P)
        End Select
    Next
    
    MakeJSONArrayFromParams = "[" & Join(Tmp, ",") & "]"

End Function

