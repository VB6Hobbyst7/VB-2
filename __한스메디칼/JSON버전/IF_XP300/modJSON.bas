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
    '-- ���ڰ� ó����
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
'''    '-- ���ڰ� ó����
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
        '-- �α���
        Case "LOGIN"
        '-- ��ũ��ȸ
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
            
        '���ڵ���ȸ
        Case "PATLIST"
            strURL = gURL.PATLIST
            strHeader = "srchMap"
    
            ReDim Preserve varPara(0) As Variant
            varPara(0) = "brcdLablNo"
    
            ReDim Preserve varVal(0) As Variant
            varVal(0) = strParam(0)
        
        '�������
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


Public Function JsonSend_EDEMIS(strAction As String, strParam() As Variant) As Variant
    Dim strURL      As String
    Dim strHeader   As String
    Dim varPara()   As Variant
    Dim varVal()    As Variant
    Dim strVHDV     As String
    Dim intAct      As Integer
    
    Select Case strAction
        '-- �α���
        Case "LOGIN"
            intAct = 1
            strURL = gURL.LOGIN
            strHeader = "srchMap"
    
            ReDim Preserve varPara(0) As Variant
            varPara(0) = "USER_ID"
    
            ReDim Preserve varVal(0) As Variant
            varVal(0) = strParam(0)
            
        '-- ��ũ��ȸ
        Case "WORKLIST"
            intAct = 2
            strURL = gURL.WORKLIST
            strHeader = "srchMap"
    
            ReDim Preserve varPara(4) As Variant
            varPara(0) = "SES_HSPT_CD"
            varPara(1) = "FROM_DATE"
            varPara(2) = "TO_DATE"
            varPara(3) = "LLRG_CD"
            varPara(4) = "VHDV_CD"
    
            ReDim Preserve varVal(4) As Variant
            varVal(0) = strParam(0)
            varVal(1) = strParam(1)
            varVal(2) = strParam(2)
            varVal(3) = strParam(3)
            varVal(4) = strParam(4)
            
        '���ڵ���ȸ
        Case "PATLIST"
            intAct = 3
            strURL = gURL.PATLIST
            strHeader = "srchMap"
    
            ReDim Preserve varPara(1) As Variant
            varPara(0) = "SES_HSPT_CD"
            varPara(1) = "BARCDNO"
    
            ReDim Preserve varVal(1) As Variant
            varVal(0) = strParam(0)
            varVal(1) = strParam(1)
        
        '�������
        Case "PATSAVE"
            intAct = 4
            strURL = gURL.PATSAVE
            strHeader = "saveList"
            
            ReDim Preserve varPara(11) As Variant
            varPara(0) = "SES_HSPT_CD"
            varPara(1) = "BARCDNO"
            varPara(2) = "PRSCRT_CODENO"
            varPara(3) = "INSP_CLSFCT_CODENO"
            varPara(4) = "SMPORE_CD"
            varPara(5) = "INSP_EQP_YN"
            varPara(6) = "INSP_EQP_CODENO"
            varPara(7) = "RLTY_RSLT_CTNT"
            varPara(8) = "APLY_RSLT_CTNT"
            varPara(9) = "RSLT_STATE_CD"
            varPara(10) = "SES_USER_ID"
            varPara(11) = "SES_USER_IP"
    
            ReDim Preserve varVal(11) As Variant
            varVal(0) = strParam(0)
            varVal(1) = strParam(1)
            varVal(2) = strParam(2)    'ó���ڵ�
            varVal(3) = strParam(3)    '�˻��ڵ�
            varVal(4) = strParam(4)    '��ü�ڵ�(�Һ�)
            varVal(5) = strParam(5)
            varVal(6) = strParam(6)    '����ڵ�
            varVal(7) = strParam(7)    '���
            varVal(8) = strParam(8)    '���
            varVal(9) = strParam(9)
            varVal(10) = strParam(10)
            varVal(11) = strParam(11)
    End Select
    
    JsonSend_EDEMIS = JSONRPC_EDEMIS(strURL, strHeader, varPara, varVal, -1)
    
    '�׽�Ʈ : �����Ͻ� �� ���� !!!!!!!!!!!!!!!!!
    JsonSend_EDEMIS = JSONRPC_EDEMIS(strURL, strHeader, varPara, varVal, intAct)
    
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

Public Function JSONRPC_EDEMIS(URL$, JSONPostHeader$, P() As Variant, V() As Variant, Optional intAct As Integer) As String
    Dim http    As Object
    Dim i       As Integer
    Dim JSONPostBody$()
  
On Error GoTo RST
    
    If intAct = -1 Then
        Set http = CreateObject("Winhttp.WinHttpRequest.5.1")
          
        http.Open "POST", URL, False
        http.setRequestHeader "Content-Type", "application/json; charset=UTF-8"
        http.setRequestHeader "Accept", "application/json"
        
        ReDim Preserve JSONPostBody$(UBound(P))
        For i = 0 To UBound(P)
            JSONPostBody(i) = MakeJSONFromParams(P(i)) & ":" & MakeJSONFromParams(V(i))
        Next
        
         If JSONPostHeader = "saveList" Then
            If JSONPostHeader <> "" Then
                http.send "{" & MakeJSONFromParams(JSONPostHeader) & ":" & "[" & "{" & Join(JSONPostBody, ",") & "}" & "]" & "}"
                Call SetRawData("[saveList]" & "{" & MakeJSONFromParams(JSONPostHeader) & ":" & "[" & "{" & Join(JSONPostBody, ",") & "}" & "]" & "}")
            Else
                http.send "{" & Join(JSONPostBody, ",") & "}"
                Call SetRawData("[saveList]" & "{" & Join(JSONPostBody, ",") & "}")
            End If
        Else
            'login, ���ڵ���ȸ, ��ũ��ȸ
            If JSONPostHeader <> "" Then
                http.send "{" & MakeJSONFromParams(JSONPostHeader) & ":" & "{" & Join(JSONPostBody, ",") & "}" & "}"
            Else
                http.send "{" & Join(JSONPostBody, ",") & "}"
            End If
        End If
        
        JSONRPC_EDEMIS = http.responseText
        Call SetRawData("[����]" & JSONRPC_EDEMIS)
        Set http = Nothing
    Else
        '=============== �׽�Ʈ �� ===============
        '-- �������ϸ�� ��θ� �����Ѵ�.
        Dim strPath     As String
        Dim strBuffer   As String
        Dim TextLine
        
        strBuffer = ""
        If intAct = 1 Then
            strPath = App.PATH & "\JSON_LOG\login.txt"
        ElseIf intAct = 2 Then
            strPath = App.PATH & "\JSON_LOG\work.txt"
        'ElseIf intAct = 3 Then
        '    strPath = App.PATH & "\JSON_LOG\barcode1.txt"
        ElseIf intAct = 3 Then
            strPath = App.PATH & "\JSON_LOG\barcode10.txt"
        ElseIf intAct = 5 Then
            strPath = App.PATH & "\JSON_LOG\save.txt"
        End If
        
        Open strPath For Input As #11 ' ������ ���ϴ�.
    
        Do While Not EOF(11) ' ������ ���� ���� ������ �ݺ��մϴ�.
            Line Input #11, TextLine ' ������ ������ ���� �о���Դϴ�.
            strBuffer = strBuffer & TextLine
        Loop
    
        Close #11 ' ������ �ݽ��ϴ�
    
        JSONRPC_EDEMIS = strBuffer
        '=============== �׽�Ʈ �� ===============
    End If

Exit Function
RST:
    Call SetRawData("[Err]" & Err.Description)
    JSONRPC_EDEMIS = ""
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

