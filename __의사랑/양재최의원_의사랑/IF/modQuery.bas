Attribute VB_Name = "modQuery"
Option Explicit

Public SQL              As String
Public RS               As ADODB.Recordset
Public blnSameRecord    As Boolean


Public Function GetEquipExamCode_STAGO(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim strExamCode     As String
    Dim strSendCH       As String
    
    GetEquipExamCode_STAGO = ""
    strExamCode = ""

    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If

    '-- 가져온 검사코드의 채널 찾기
    SQL = ""
    SQL = SQL & "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"

    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strSendCH = Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
            If strSendCH <> "" Then
                strExamCode = strExamCode & "^^^" & strSendCH & "^\"
            End If
            
            If strSendCH = "2" Then
                strExamCode = strExamCode & "\^^^1"
            End If
            
            If strSendCH <> "990" Then
                strExamCode = strExamCode & "\^^^" & strSendCH
            End If
            
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close
    
    If strExamCode <> "" Then
        GetEquipExamCode_STAGO = Mid(strExamCode, 2)
    End If
    
End Function

Public Function GetEquipExamCode_ACCESS2(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim strExamCode     As String
    Dim strSendCH       As String
    
    GetEquipExamCode_ACCESS2 = ""
    strExamCode = ""
    mOrder.SendCnt = 0
    
    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If

    '-- 가져온 검사코드의 채널 찾기
    SQL = ""
    SQL = SQL & "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"

    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strSendCH = Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
            If strSendCH <> "" Then
                mOrder.SendCnt = mOrder.SendCnt + 1
                If strExamCode = "" Then
                    strExamCode = "^^^" & strSendCH & "^1"
                Else
                    strExamCode = strExamCode & "\^^^" & strSendCH & "^1"
                End If
            End If
            
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close
    
    If strExamCode <> "" Then
        GetEquipExamCode_ACCESS2 = strExamCode
    End If
    
End Function


Public Function GetEquipExamCode_AU480(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim strExamCode     As String
    Dim strSendCH       As String
    
    GetEquipExamCode_AU480 = ""
    strExamCode = ""
    mOrder.SendCnt = 0
    
    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If

    '-- 가져온 검사코드의 채널 찾기
    SQL = ""
    SQL = SQL & "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"

    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strSendCH = Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
            If strSendCH <> "" Then
                mOrder.SendCnt = mOrder.SendCnt + 1
                strExamCode = strExamCode & Format(strSendCH, "000") & "0"
            End If
            
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close
    
    If strExamCode <> "" Then
        GetEquipExamCode_AU480 = strExamCode
    End If
    
End Function

Public Function GetEquipExamCode_PPC300N(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim strExamCode     As String
    Dim strSendCH       As String
    
    GetEquipExamCode_PPC300N = ""
    strExamCode = ""

    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If

    '-- 가져온 검사코드의 채널 찾기
    SQL = ""
    SQL = SQL & "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"

    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    
    'AST^ALT^TP^GLU_HK
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strSendCH = Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
            If strSendCH <> "" Then
                If strExamCode = "" Then
                    strExamCode = strSendCH
                Else
                    strExamCode = strExamCode & "^" & strSendCH
                End If
            End If
            
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close
    
    If strExamCode <> "" Then
        GetEquipExamCode_PPC300N = strExamCode
    End If
    
End Function

Public Function GetEquipExamCode_ACCESS(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim strExamCode     As String
    Dim strSendCH       As String
    
    GetEquipExamCode_ACCESS = ""
    strExamCode = ""

    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If

    '-- 가져온 검사코드의 채널 찾기
    SQL = ""
    SQL = SQL & "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"

    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strSendCH = Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
            If strSendCH <> "" Then
                strExamCode = strExamCode & "^^^" & strSendCH & "\"
                mOrder.SendCnt = mOrder.SendCnt + 1
            End If
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close
    
    If strExamCode <> "" Then
        GetEquipExamCode_ACCESS = Mid(strExamCode, 1, Len(strExamCode) - 1)
    End If
    
End Function


Public Function GetEquipExamCode_C8000(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim strExamCode     As String
    Dim strSendCH       As String
    Dim iAffected       As Integer
    
    GetEquipExamCode_C8000 = ""
    strExamCode = ""

    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If

    '-- 가져온 검사코드의 채널 찾기
    SQL = ""
    SQL = SQL & "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"

    AdoCn_Local.CursorLocation = adUseClient
    
    'CommandText        String  수행할 명령을 기술하는 매개변수이며, SQL 문장, 테이블 명, 저장 프로시저를 지정할 수 있다.
    'RecordsAffected    Long    Execute 메서드에 의해서 영향을 받은 레코드의 개수를 반환한다. 예를 들면 Delete문장을 수행했는데, 10 개의 레코드가 삭제되었다면, 10 이라는 값을 반환한다.
    'Options            Long    Provider가 CommandText를 어떻게 수행할지를 결정하는 방법을 지정하는 값이며, 데이터 형식은 Long이다.
    '                    1      : adCmdText         CommandText의 값을 SQL 문장으로 처리한다.
    '                    2      : adCmdTable        CommandText의 값을 테이블 명으로 하는 SQL 문장을 만들어서 처리한다.
    '                    512    : adCmdTableDirect  CommandText의 값을 테이블 명으로 처리한다.
    '                    4      : adCmdStoredProc   CommandText의 값을 저장 프로시저로 처리한다.
    '                    8      : adCmdUnknown      명령의 형식을 알 수 없음으로 처리한다.
    '                    16     : adAsyncExecute    명령을 비동기적으로 수행한다.
    '                    32     : adAsyncFetch      CacheSize 속성에 지정된 수 만큼의 레코드씩 비동기적으로 처리한다.
    
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, iAffected, 1)
    
    'iAffected = AdoRs_Local.RecordCount
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strSendCH = Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
            If strSendCH <> "" Then
                strExamCode = strExamCode & "\^^^" & strSendCH & "^1"
            End If
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close
    
    If strExamCode <> "" Then
        GetEquipExamCode_C8000 = Mid(strExamCode, 2)
    End If
    
End Function


Public Function GetEquipExamCode_CA800(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim strExamCode     As String
    Dim strSendCH       As String
    Dim iAffected       As Integer
    
    GetEquipExamCode_CA800 = ""
    strExamCode = ""

    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If

    '-- 가져온 검사코드의 채널 찾기
    SQL = ""
    SQL = SQL & "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"

    AdoCn_Local.CursorLocation = adUseClient
    
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, iAffected, 1)
    '^^^040^^100\^^^050^^100
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strSendCH = Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
            If strSendCH <> "" Then
                strSendCH = Format(strSendCH, "000")
                strExamCode = strExamCode & strSendCH & Space(6)
            End If
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close
    
    If strExamCode <> "" Then
        GetEquipExamCode_CA800 = strExamCode
    End If
    
End Function

Public Function GetEquipExamCode_E411(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim strExamCode     As String
    Dim strSendCH       As String
    Dim iAffected       As Integer
    
    GetEquipExamCode_E411 = ""
    strExamCode = ""

    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If

    '-- 가져온 검사코드의 채널 찾기
    SQL = ""
    SQL = SQL & "Select DISTINCT a.SENDCHANNEL as SENDCHANNEL" & vbCrLf
    SQL = SQL & "  From EQPMASTER a, TESTMASTER b " & vbCrLf
    SQL = SQL & " Where a.RSLTCHANNEL = b.RSLTCHANNEL " & vbCrLf
    If gPatOrdCd <> "" Then
        SQL = SQL & "   And b.TESTCODE IN (" & Trim(gPatOrdCd) & ")" & vbCrLf
    End If
    SQL = SQL & "   And (a.SENDCHANNEL <> '' OR a.SENDCHANNEL IS NOT NULL)" & vbCrLf

    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, iAffected, 1)
    '^^^301^\^^^311^
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strSendCH = Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
            If strSendCH <> "" Then
                strExamCode = strExamCode & "\^^^" & strSendCH & "^"
            End If
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close
    
    If strExamCode <> "" Then
        GetEquipExamCode_E411 = Mid(strExamCode, 2)
    End If
    
End Function

Public Function GetEquipExamCode_CA800_ASTM(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim strExamCode     As String
    Dim strSendCH       As String
    Dim iAffected       As Integer
    
    GetEquipExamCode_CA800_ASTM = ""
    strExamCode = ""

    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If

    '-- 가져온 검사코드의 채널 찾기
    SQL = ""
    SQL = SQL & "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"

    AdoCn_Local.CursorLocation = adUseClient
    
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, iAffected, 1)
    '^^^040^^100\^^^050^^100
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strSendCH = Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
            If strSendCH <> "" Then
                strExamCode = strExamCode & "\^^^" & strSendCH & "^^100"
            End If
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close
    
    If strExamCode <> "" Then
        GetEquipExamCode_CA800_ASTM = Mid(strExamCode, 2)
    End If
    
End Function

Public Function GetEquipExamCode_XN1000(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim strExamCode     As String
    Dim strSendCH       As String
    Dim strCBC      As String
    Dim strDIFF     As String
    
    GetEquipExamCode_XN1000 = ""
    strExamCode = ""

    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        '-- 오더가 없을 경우 CBC/ DIFF 검사하도록 한다.
        If strExamCode = "" Then
            strExamCode = "^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\^^^^RDW-SD\^^^^RDW-CV\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT\^^^^NRBC#\^^^^NRBC%\"
            'strExamCode = strExamCode & "^^^^NEUT#\^^^^LYMPH%\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH#\^^^^LYMPH#\^^^^MONO%\^^^^EO%\^^^^BASO%\^^^^IG#\^^^^IG%\"
            strExamCode = strExamCode & "^^^^NEUT#\^^^^LYMPH%\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH#\^^^^MONO%\^^^^EO%\^^^^BASO%\^^^^IG#\^^^^IG%\"
        End If
        
        If strExamCode <> "" Then
            strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
        End If
        
        GetEquipExamCode_XN1000 = strExamCode
        
        Exit Function
    End If

    '-- 가져온 검사코드의 채널 찾기
    SQL = ""
    SQL = SQL & "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"

    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strSendCH = Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
            If strSendCH <> "" Then
                If strSendCH = "WBC" Or strSendCH = "RBC" Or strSendCH = "HGB" Or _
                    strSendCH = "HCT" Or strSendCH = "MCV" Or strSendCH = "MCH" Or strSendCH = "MCHC" Or _
                    strSendCH = "PLT" Or strSendCH = "RDW-SD" Or strSendCH = "RDW-CV" Or strSendCH = "PDW" Or _
                    strSendCH = "MPV" Or strSendCH = "P-LCR" Or strSendCH = "PCT" Or strSendCH = "NRBC#" Or strSendCH = "NRBC%" Then
                    
                    strCBC = "^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\^^^^RDW-SD\^^^^RDW-CV\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT\^^^^NRBC#\^^^^NRBC%\"
                    
                End If
    
                If strSendCH = "NEUT#" Or strSendCH = "LYMPH#" Or strSendCH = "MONO#" Or strSendCH = "EO#" Or strSendCH = "BASO#" Or _
                    strSendCH = "NEUT%" Or strSendCH = "LYMPH%" Or strSendCH = "MONO%" Or strSendCH = "EO%" Or strSendCH = "BASO%" Or _
                    strSendCH = "IG#" Or strSendCH = "IG%" Then
                   
                    '-- ^^^^LYMPH#\가 두개인 이유는 ETB 를 장비에서 인식하지 못하기 떄문..(그 자리가 230)
                    'strDIFF = "^^^^NEUT#\^^^^LYMPH%\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH#\^^^^LYMPH#\^^^^MONO%\^^^^EO%\^^^^BASO%\^^^^IG#\^^^^IG%\"
                    strDIFF = "^^^^NEUT#\^^^^LYMPH%\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH#\^^^^MONO%\^^^^EO%\^^^^BASO%\^^^^IG#\^^^^IG%\"
                    
                End If
            End If
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close
    
    strExamCode = strCBC & strDIFF
    
    '-- 오더가 없을 경우 CBC/ DIFF 검사하도록 한다.
    If strExamCode = "" Then
        strExamCode = "^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\^^^^RDW-SD\^^^^RDW-CV\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT\^^^^NRBC#\^^^^NRBC%\"
        'strExamCode = strExamCode & "^^^^NEUT#\^^^^LYMPH%\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH#\^^^^LYMPH#\^^^^MONO%\^^^^EO%\^^^^BASO%\^^^^IG#\^^^^IG%\"
        strExamCode = strExamCode & "^^^^NEUT#\^^^^LYMPH%\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH#\^^^^MONO%\^^^^EO%\^^^^BASO%\^^^^IG#\^^^^IG%\"
    End If
    
    If strExamCode <> "" Then
        GetEquipExamCode_XN1000 = Mid(strExamCode, 1, Len(strExamCode) - 1)
    End If
    
End Function

Public Function GetEquipExamCode_THUNDERBOLT(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim sExamCode As String
    Dim strExamCode As String
    Dim sSpecNo     As String
    Dim iRow        As Long
    Dim SpecNo      As String

    GetEquipExamCode_THUNDERBOLT = ""

    'strExamCode = "@^^^Measles_IgG@^^^Mumps_IgG@^^^VZV_IgG"
    'strExamCode = "@^^^ADENO"
    'GetEquipExamCode_THUNDERBOLT = Mid(strExamCode, 2)

    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If

    '-- 가져온 검사코드의 채널 찾기
          SQL = "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"

    strExamCode = ""

    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strExamCode = strExamCode & "@^^^" & Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
            AdoRs_Local.MoveNext
        Loop
    End If
    
'    strExamCode = "@^^^Measles_IgG^^^Mumps_IgG@^^^VZV_IgG"

    AdoRs_Local.Close

    GetEquipExamCode_THUNDERBOLT = Mid(strExamCode, 2)

End Function

Public Function GetEquipExamCode_YUMIZEN(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim strExamCode     As String
    Dim strSendCH       As String
    Dim strCBC      As String
    Dim strDIFF     As String
    
    GetEquipExamCode_YUMIZEN = ""
    strExamCode = ""
    mOrder.SendCnt = 0

    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        '-- 오더가 없을 경우 CBC/DIFF 검사하도록 한다.
        If strExamCode = "" Then
            strExamCode = "^^^DIF"
        End If
        
        GetEquipExamCode_YUMIZEN = strExamCode
        
        Exit Function
    End If

    '-- 가져온 검사코드의 채널 찾기
    SQL = ""
    SQL = SQL & "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"

    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strSendCH = Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
            If strSendCH <> "" Then
                If strSendCH = "WBC" Or strSendCH = "RBC" Or strSendCH = "HGB" Or _
                    strSendCH = "HCT" Or strSendCH = "MCV" Or strSendCH = "MCH" Or strSendCH = "MCHC" Or _
                    strSendCH = "PLT" Or strSendCH = "RDW-SD" Or strSendCH = "RDW-CV" Or strSendCH = "PDW" Or _
                    strSendCH = "MPV" Or strSendCH = "P-LCR" Or strSendCH = "PCT" Or strSendCH = "NRBC#" Or strSendCH = "NRBC%" Then
                    
                    strCBC = "^^^CBC"
                    mOrder.SendCnt = 1
                End If
    
                If strSendCH = "NEU#" Or strSendCH = "LYM#" Or strSendCH = "MON#" Or strSendCH = "EOS#" Or strSendCH = "BAS#" Or _
                    strSendCH = "NEU%" Or strSendCH = "LYM%" Or strSendCH = "MON%" Or strSendCH = "EOS%" Or strSendCH = "BAS%" Then
                   
                    strDIFF = "^^^DIF"
                    mOrder.SendCnt = 1
                End If
            End If
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close
    
    If strCBC <> "" Then
        strExamCode = strCBC
    End If
    
    If strDIFF <> "" Then
        strExamCode = strDIFF
    End If
    
    '-- 오더가 없을 경우 CBC/ DIFF 검사하도록 한다.
    If strExamCode = "" Then
        strExamCode = "^^^DIF"
    End If
    
    If strExamCode <> "" Then
        GetEquipExamCode_YUMIZEN = strExamCode
    End If
    
End Function


'한 장비채널에 검사코드가 1개이상 존재 (GLU-FBS, GLU-PP2..)
Public Function GetEquipExamCode_AU680(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim strExamCode     As String
    Dim strSendCH       As String
    
    GetEquipExamCode_AU680 = ""
    strExamCode = ""

    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If

    '-- 가져온 검사코드의 채널 찾기
    SQL = ""
    SQL = SQL & "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"

    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strSendCH = Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
            If strSendCH <> "" Then
                strExamCode = strExamCode & Format(strSendCH, "000")
            End If
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close

    GetEquipExamCode_AU680 = strExamCode

End Function

'한 장비채널에 검사코드가 1개이상 존재 (GLU-FBS, GLU-PP2..)
Public Function GetEquipExamCode_HITACHI7180(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim strExamCode     As String
    Dim intIntBase      As Integer
    Dim strItems        As String           '전송할 검사항목
    Dim blnISE          As Boolean          'Na, K, Cl 검사여부

    strItems = String$(88, "0")
    
    GetEquipExamCode_HITACHI7180 = strItems
    strExamCode = ""
    blnISE = False
    mOrder.SendCnt = 0
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If

    '-- 가져온 검사코드의 채널 찾기
    SQL = ""
    SQL = SQL & "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    If gPatOrdCd <> "" Then
        SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"
    Else
        GetEquipExamCode_HITACHI7180 = strItems
        Exit Function
    End If
    
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            If AdoRs_Local.Fields("SENDCHANNEL").Value & "" <> "" And IsNumeric(AdoRs_Local.Fields("SENDCHANNEL").Value & "") Then
                intIntBase = AdoRs_Local.Fields("SENDCHANNEL").Value & ""
                If intIntBase = "99" Then 'GA%
                    'GA
                    Mid$(strItems, 25, 1) = "1"
                    'GA-Alb
                    Mid$(strItems, 26, 1) = "1"
                Else
                    Mid$(strItems, intIntBase, 1) = "1"
                End If
                mOrder.SendCnt = mOrder.SendCnt + 1
            End If
'
'            If IsNumeric(AdoRs_Local.Fields("SENDCHANNEL").Value) Then
'
'                intIntBase = CInt(AdoRs_Local.Fields("SENDCHANNEL").Value)
'                If intIntBase <> "" Then
'                    '## 계산항목: 93~100
'                    If intIntBase >= 93 And intIntBase <= 100 Then
'                        'GoTo Skip1
'                    Else
'                        '## Na, K, Cl 검사여부 Check
'                        If intIntBase = 87 Or intIntBase = 88 Or intIntBase = 89 Then
'                            blnISE = True
'                        Else
'                            Mid$(strItems, intIntBase, 1) = "1"
'                        End If
'                    End If
'                    mOrder.SendCnt = mOrder.SendCnt + 1
'                End If
'            End If
            
            AdoRs_Local.MoveNext
        Loop
    End If

    '## Na, K, Cl 검사여부 Check
'    If blnISE Then
'        Mid$(strItems, 87, 1) = "1"
'        mOrder.SendCnt = mOrder.SendCnt + 1
'    End If

    AdoRs_Local.Close

    'Call SetSQLData("strItems", strItems)

    GetEquipExamCode_HITACHI7180 = strItems
    
  '  MsgBox strItems

End Function

'한 장비채널에 검사코드가 1개이상 존재 (GLU-FBS, GLU-PP2..)
Public Function GetEquipExamCode_HITACHI7020(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim strExamCode     As String
    Dim intIntBase      As Integer
    Dim strItems        As String           '전송할 검사항목
    'Dim blnISE          As Boolean          'Na, K, Cl 검사여부

    strItems = String$(37, "0")
    
    GetEquipExamCode_HITACHI7020 = strItems
    strExamCode = ""
    'blnISE = False
    mOrder.SendCnt = 0
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If

    '-- 가져온 검사코드의 채널 찾기
    SQL = ""
    SQL = SQL & "Select DISTINCT a.SENDCHANNEL as SENDCHANNEL" & vbCrLf
    SQL = SQL & "  From EQPMASTER a, TESTMASTER b " & vbCrLf
    SQL = SQL & " Where a.RSLTCHANNEL = b.RSLTCHANNEL " & vbCrLf
    If gPatOrdCd <> "" Then
        SQL = SQL & "   And b.TESTCODE IN (" & Trim(gPatOrdCd) & ")" & vbCrLf
    End If
    SQL = SQL & "   And (a.SENDCHANNEL <> '' OR a.SENDCHANNEL IS NOT NULL)" & vbCrLf

    'Call SetSQLData("채널조회", SQL, "A")

    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)

    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            If AdoRs_Local.Fields("SENDCHANNEL").Value & "" <> "" Then
                intIntBase = Val(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
                If intIntBase <= 37 Then
                    If intIntBase <> 99 Then
                        Mid$(strItems, intIntBase, 1) = "1"
                        mOrder.SendCnt = mOrder.SendCnt + 1
                    End If
                End If
            End If
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close

    GetEquipExamCode_HITACHI7020 = strItems
    
'    Call SetSQLData("바코드조회", ">>>" & GetEquipExamCode_HITACHI7020, "A")
    
End Function

Function SaveTransData_EONM(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    Dim RsLocal         As ADODB.Recordset
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strPatNm        As String
    
    Dim strEqpcd        As String
    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strTestCdSub    As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strJudge        As String
    
On Error GoTo ErrHandle
    
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""

    With frmInterface
        SaveTransData_EONM = -1
        
        strSaveSeq = Trim(GetText(SPD, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(SPD, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(SPD, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(SPD, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(SPD, argSpcRow, colPID))
        strPatNm = Trim(GetText(SPD, argSpcRow, colPNAME))
        strChartNo = Trim(GetText(SPD, argSpcRow, colCHARTNO))
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatNm) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMCODESUB,EQUIPRESULT,RESULT,REFJUDGE    " & vbCrLf
        SQL = SQL & "  FROM PATRESULT                                                               " & vbCrLf
        SQL = SQL & " WHERE EXAMDATE    = '" & strExamDate & "'                                     " & vbCrLf
        SQL = SQL & "   AND SAVESEQ     = " & strSaveSeq & vbCrLf
        SQL = SQL & "   AND BARCODE     = '" & strBarcode & "'                                      " & vbCrLf
        SQL = SQL & "   AND EXAMCODE    <> ''                                                       " & vbCrLf
        
        Set RsLocal = New ADODB.Recordset
        Set RsLocal = AdoCn_Local.Execute(SQL, , 1)
        If Not RsLocal.EOF = True And Not RsLocal.BOF = True Then
            Do Until RsLocal.EOF
                strEqpcd = RsLocal.Fields("EQUIPCODE").Value & ""
                strOrdCd = RsLocal.Fields("ORDERCODE").Value & ""
                strTestCd = RsLocal.Fields("EXAMCODE").Value & ""
                strTestCdSub = RsLocal.Fields("EXAMCODESUB").Value & ""
                sResult1 = RsLocal.Fields("EQUIPRESULT").Value & ""
                sResult2 = RsLocal.Fields("RESULT").Value & ""
                strJudge = RsLocal.Fields("REFJUDGE").Value & ""
                
                '-- 장비결과적용
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                  
                '-- 결과저장용 키 가져오기
                strTestCdSub = GetSampleSubITEM(strBarcode, strTestCd)

'                SQL = ""
'                SQL = SQL & "SELECT DISTINCT O.H141_SEQNO   AS SUBITEM    " & vbCr
'                SQL = SQL & "  FROM TB_H141_LISTAKEBODY O, TB_A110_PATINFO P    " & vbCr
'                SQL = SQL & " Where P.A110_ChartNo    = O.H141_CHARTNO             " & vbCr
'                SQL = SQL & "   AND O.H141_TSAMPLENO  = '" & strBarcode & "'    " & vbCr
'                SQL = SQL & "   And O.H141_SUGACD     = '" & strTestCd & "'    " & vbCrLf
'
'                Set RS = AdoCn.Execute(SQL, , 1)
'                If Not RS.EOF = True And Not RS.BOF = True Then
'                    Do Until RS.EOF
'                        strTestCdSub = Trim(RS.Fields("SUBITEM")) & ""
'                        RS.MoveNext
'                    Loop
'                End If
                
                RS.Close
                
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" And strTestCdSub <> "" Then
                    '-- 서버저장
                    SQL = "" '
                    SQL = SQL & "Update TB_H141_LISTAKEBODY                     " & vbCrLf
                    SQL = SQL & "   SET H141_RSLTYN    ='Y'                     " & vbCrLf
                    SQL = SQL & " WHERE H141_TSAMPLENO = '" & strBarcode & "'   " & vbCrLf
                    SQL = SQL & "   AND H141_SUGACD    = '" & strTestCd & "'    " & vbCrLf
                    
                    Call SetSQLData("결과저장", SQL, "A")
                    AdoCn.Execute SQL
                    
                    SQL = ""
                    SQL = SQL & "UPDATE TB_H131_SPPRESULT                       " & vbCrLf
                    SQL = SQL & "   SET H131_RESULT  = '" & sResult & "'        " & vbCrLf
                    SQL = SQL & " WHERE H131_SPPTYPE = '" & gHOSP.PARTCD & "'   " & vbCrLf    'L010
                    SQL = SQL & "   AND H131_SEQNO   = '" & strTestCdSub & "'   " & vbCrLf
                        
                    Call SetSQLData("결과저장", SQL, "A")
                    AdoCn.Execute SQL
                
                    SQL = ""
                    SQL = SQL & "UPDATE TB_H130_SPPRECEIVE                              " & vbCrLf
                    SQL = SQL & "   SET H130_RSLTDAT = TO_CHAR(SYSDATE, 'YYYYMMDD')     " & vbCrLf
                    SQL = SQL & "      ,H130_RSLTTM  = TO_CHAR(SYSDATE, 'HH24:MI:SS')   " & vbCrLf
                    SQL = SQL & " WHERE H130_SPPTYPE = '" & gHOSP.PARTCD & "'           " & vbCrLf    'L010
                    SQL = SQL & "   AND H130_SEQNO   = '" & strTestCdSub & "'           " & vbCrLf
                        
                    Call SetSQLData("결과저장", SQL, "A")
                    AdoCn.Execute SQL
                
                    SQL = ""
                    SQL = SQL & "UPDATE TB_H140_LISTAKEHEAD                     " & vbCrLf
                    SQL = SQL & "   SET H140_RSLTYN    = 'Y'                    " & vbCrLf
                    SQL = SQL & " WHERE H140_TSAMPLENO = '" & strBarcode & "'   " & vbCrLf
                                        
                    Call SetSQLData("결과저장", SQL, "A")
                    AdoCn.Execute SQL
                            
                End If
                RsLocal.MoveNext
            Loop
        End If
        
        RsLocal.Close
        
        SaveTransData_EONM = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_EONM = -1
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "SaveTransData_EONM" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

Function SaveTransData_AMIS(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    Dim RsLocal         As ADODB.Recordset
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strPatNm        As String
    
    Dim strEqpcd        As String
    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strTestCdSub    As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strJudge        As String
    
On Error GoTo ErrHandle
    
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""

    With frmInterface
        SaveTransData_AMIS = -1
        
        strSaveSeq = Trim(GetText(SPD, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(SPD, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(SPD, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(SPD, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(SPD, argSpcRow, colPID))
        strPatNm = Trim(GetText(SPD, argSpcRow, colPNAME))
        strChartNo = Trim(GetText(SPD, argSpcRow, colCHARTNO))
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatNm) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMCODESUB,EQUIPRESULT,RESULT,REFJUDGE    " & vbCrLf
        SQL = SQL & "  FROM PATRESULT                                                               " & vbCrLf
        SQL = SQL & " WHERE EXAMDATE    = '" & strExamDate & "'                                     " & vbCrLf
        SQL = SQL & "   AND SAVESEQ     = " & strSaveSeq & vbCrLf
        SQL = SQL & "   AND BARCODE     = '" & strBarcode & "'                                      " & vbCrLf
        SQL = SQL & "   AND EXAMCODE    <> ''                                                       " & vbCrLf
        
        Set RsLocal = New ADODB.Recordset
        Set RsLocal = AdoCn_Local.Execute(SQL, , 1)
        If Not RsLocal.EOF = True And Not RsLocal.BOF = True Then
            Do Until RsLocal.EOF
                strEqpcd = RsLocal.Fields("EQUIPCODE").Value & ""
                strOrdCd = RsLocal.Fields("ORDERCODE").Value & ""
                strTestCd = RsLocal.Fields("EXAMCODE").Value & ""
                strTestCdSub = RsLocal.Fields("EXAMCODESUB").Value & ""
                sResult1 = RsLocal.Fields("EQUIPRESULT").Value & ""
                sResult2 = RsLocal.Fields("RESULT").Value & ""
                strJudge = RsLocal.Fields("REFJUDGE").Value & ""
                
                '-- 장비결과적용
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                  
                '-- 결과저장용 키 가져오기
                If strOrdCd = "" Then
                    strOrdCd = GetSampleSubITEM(strBarcode, strTestCd)
                End If
                
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" And strOrdCd <> "" Then
                    '-- 서버저장
                    SQL = ""
                    SQL = SQL & "Update RESULTOFNUM                                     " & vbCrLf
                    SQL = SQL & "   Set RESULTINDATE   = to_char(sysdate,'yyyymmdd')    " & vbCrLf
                    SQL = SQL & "     , RESULTINTIME   = to_char(sysdate,'HH24MI')      " & vbCrLf
                    SQL = SQL & "     , RESULTINID     = '" & gHOSP.USERID & "'         " & vbCrLf
                    SQL = SQL & "     , RESULTFLAG     = '1'                            " & vbCrLf
                    SQL = SQL & "     , TEXTRESULTVAL  = '" & sResult & "'              " & vbCrLf
                    '-- 결과가 수치형이면
                    If IsNumeric(sResult) Then
                        SQL = SQL & "     , NUMRESULTVAL = '" & sResult & "'           " & vbCrLf
                    End If
                    SQL = SQL & " Where SPCMNO         = '" & strBarcode & "'           " & vbCrLf
                    SQL = SQL & "   And ORDERCODE      = '" & strOrdCd & "'             " & vbCrLf
                    SQL = SQL & "   And RESULTITEMCODE = '" & strTestCd & "'            " & vbCrLf
                    SQL = SQL & "   And RESULTFLAG < '3'                                " & vbCrLf
                    
                    Call SetSQLData("결과저장", SQL, "A")
                    AdoCn.Execute SQL
                                        
                    '-- 상태변경
                    SQL = ""
                    SQL = SQL & "Update REGISTINFOS                         " & vbCrLf
                    SQL = SQL & "   Set RESULTSTATE  = '1'                  " & vbCrLf
                    SQL = SQL & "      ,RsvAcptState = '4'                  " & vbCrLf
                    SQL = SQL & " Where SPCMNO       = '" & strBarcode & "' " & vbCrLf
                    SQL = SQL & "   AND ORDERCODE    = '" & strOrdCd & "'   " & vbCrLf
                    SQL = SQL & "   AND CLAS         = 4                    " & vbCrLf
                    SQL = SQL & "   AND RESULTSTATE < '4'                   " & vbCrLf
                    
                    Call SetSQLData("상태변경", SQL, "A")
                    AdoCn.Execute SQL
                            
                End If
                RsLocal.MoveNext
            Loop
        End If
        
        RsLocal.Close
        
        SaveTransData_AMIS = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_AMIS = -1
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SaveTransData_AMIS" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

Function SaveTransData_KCHART(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    Dim RsLocal         As ADODB.Recordset
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strPatNm        As String
    
    Dim strEqpcd        As String
    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strTestCdSub    As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strJudge        As String
    
On Error GoTo ErrHandle
    
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""

    With frmInterface
        SaveTransData_KCHART = -1
        
        strSaveSeq = Trim(GetText(SPD, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(SPD, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(SPD, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(SPD, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(SPD, argSpcRow, colPID))
        strPatNm = Trim(GetText(SPD, argSpcRow, colPNAME))
        strChartNo = Trim(GetText(SPD, argSpcRow, colCHARTNO))
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatNm) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMCODESUB,EQUIPRESULT,RESULT,REFJUDGE    " & vbCrLf
        SQL = SQL & "  FROM PATRESULT                                                               " & vbCrLf
        SQL = SQL & " WHERE EXAMDATE    = '" & strExamDate & "'                                     " & vbCrLf
        SQL = SQL & "   AND SAVESEQ     = " & strSaveSeq & vbCrLf
        SQL = SQL & "   AND BARCODE     = '" & strBarcode & "'                                      " & vbCrLf
        SQL = SQL & "   AND EXAMCODE    <> ''                                                       " & vbCrLf
        
        Set RsLocal = New ADODB.Recordset
        Set RsLocal = AdoCn_Local.Execute(SQL, , 1)
        If Not RsLocal.EOF = True And Not RsLocal.BOF = True Then
            Do Until RsLocal.EOF
                strEqpcd = RsLocal.Fields("EQUIPCODE").Value & ""
                strOrdCd = RsLocal.Fields("ORDERCODE").Value & ""
                strTestCd = RsLocal.Fields("EXAMCODE").Value & ""
                strTestCdSub = RsLocal.Fields("EXAMCODESUB").Value & ""
                sResult1 = RsLocal.Fields("EQUIPRESULT").Value & ""
                sResult2 = RsLocal.Fields("RESULT").Value & ""
                strJudge = RsLocal.Fields("REFJUDGE").Value & ""
                
                '-- 장비결과적용
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                
                'MsgBox strOrdCd & "," & strTestCd & "," & strTestCdSub
                
                
                '-- 결과저장용 키 가져오기
                If strOrdCd = "" Then
                    strOrdCd = GetSampleSubITEM(strBarcode, strTestCd)
                    strOrdCd = mGetP(strOrdCd, 1, "|")
                    strTestCdSub = mGetP(strOrdCd, 2, "|")
                End If
                
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" And strOrdCd <> "" And strTestCdSub <> "" Then
                    '-- 결과저장
                    'SQL = SQL & "    ,  수정자 = 'IIS', " & vbCr
                    SQL = ""
                    SQL = SQL & "Update TB_진료검사                                   " & vbCrLf
                    SQL = SQL & "   Set 검사결과              = '" & sResult & "'     " & vbCrLf
                    SQL = SQL & "     , 하이로우              = '" & strJudge & "'    " & vbCrLf
                    SQL = SQL & "     , 검사상태              = '2'                   " & vbCrLf
                    SQL = SQL & "     , 연동구분              = '1'                   " & vbCrLf
                    SQL = SQL & "     , 수정일자              = GetDate()             " & vbCrLf
                    SQL = SQL & " Where 진료검사ID            = '" & strOrdCd & "'    " & vbCrLf
                    SQL = SQL & "   And 진료지원ID            = '" & strTestCdSub & "'" & vbCrLf
                    SQL = SQL & "   And 검체번호              = '" & strBarcode & "'  " & vbCrLf
                    SQL = SQL & "   And (처방코드 + 서브코드) = '" & strTestCd & "'   " & vbCrLf
                    
                    Call SetSQLData("결과저장", SQL, "A")
                    AdoCn.Execute SQL
                                        
                            
                End If
                RsLocal.MoveNext
            Loop
        End If
        
        RsLocal.Close
        
        SaveTransData_KCHART = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_KCHART = -1
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SaveTransData_KCHART" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

Function SaveTransData_JWINFO(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    Dim RsLocal         As ADODB.Recordset
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strPatNm        As String
    
    Dim strEqpcd        As String
    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strTestCdSub    As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strJudge        As String
    
On Error GoTo ErrHandle
    
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""

    With frmInterface
        SaveTransData_JWINFO = -1
        
        strSaveSeq = Trim(GetText(SPD, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(SPD, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(SPD, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(SPD, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(SPD, argSpcRow, colPID))
        strPatNm = Trim(GetText(SPD, argSpcRow, colPNAME))
        strChartNo = Trim(GetText(SPD, argSpcRow, colCHARTNO))
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatNm) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMCODESUB,EQUIPRESULT,RESULT,REFJUDGE    " & vbCrLf
        SQL = SQL & "  FROM PATRESULT                                                               " & vbCrLf
        SQL = SQL & " WHERE EXAMDATE    = '" & strExamDate & "'                                     " & vbCrLf
        SQL = SQL & "   AND SAVESEQ     = " & strSaveSeq & vbCrLf
        SQL = SQL & "   AND BARCODE     = '" & strBarcode & "'                                      " & vbCrLf
        SQL = SQL & "   AND EXAMCODE    <> ''                                                       " & vbCrLf
        
        Set RsLocal = New ADODB.Recordset
        Set RsLocal = AdoCn_Local.Execute(SQL, , 1)
        If Not RsLocal.EOF = True And Not RsLocal.BOF = True Then
            Do Until RsLocal.EOF
                strEqpcd = RsLocal.Fields("EQUIPCODE").Value & ""
                strOrdCd = RsLocal.Fields("ORDERCODE").Value & ""
                strTestCd = RsLocal.Fields("EXAMCODE").Value & ""
                strTestCdSub = RsLocal.Fields("EXAMCODESUB").Value & ""
                sResult1 = RsLocal.Fields("EQUIPRESULT").Value & ""
                sResult2 = RsLocal.Fields("RESULT").Value & ""
                strJudge = RsLocal.Fields("REFJUDGE").Value & ""
                
                '-- 장비결과적용
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                
                '-- 결과저장용 키 가져오기
                If strOrdCd = "" Then
                    strOrdCd = GetSampleSubITEM(strBarcode, strTestCd)
                End If
                
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" And strOrdCd <> "" Then
                    '-- 결과저장
                          SQL = "Update SLA_LabResult                       " & vbCrLf
                    SQL = SQL & "   Set Result      = '" & sResult & "'     " & vbCrLf
                    SQL = SQL & "      ,NormalFlag  = '0'                   " & vbCrLf
                    SQL = SQL & "      ,PanicFlag   = '0'                   " & vbCrLf
                    SQL = SQL & "      ,DeltaFlag   = '0'                   " & vbCrLf
                    SQL = SQL & "      ,TransFlag   = '1'                   " & vbCrLf
                    SQL = SQL & "      ,ResultID    = '" & gHOSP.USERID & "'" & vbCrLf
                    SQL = SQL & "      ,ResultDate  = '" & Trim(Format(Now, "yyyy-mm-dd")) & "'" & vbCrLf
                    SQL = SQL & "      ,ResultTime  = '" & Trim(Format(Time, "HH:MM:SS")) & "'" & vbCrLf
                    SQL = SQL & " Where SPECIMENNUM = '" & strBarcode & "'  " & vbCrLf
                    SQL = SQL & "   AND OrderCode   = '" & strOrdCd & "'    " & vbCrLf
                    SQL = SQL & "   And LabCode     = '" & strTestCd & "'   " & vbCrLf
                    SQL = SQL & "   And TRANSFLAG   < '2'                   " & vbCrLf
                    
                    Call SetSQLData("결과저장", SQL, "A")
                    AdoCn.Execute SQL
                    
                    '-- 상태변경
                          SQL = "Update SLA_LabMaster                       " & vbCrLf
                    SQL = SQL & "   Set JStatus = '2'                       " & vbCrLf
                    SQL = SQL & " Where SPECIMENNUM = '" & strBarcode & "'  " & vbCrLf
                    SQL = SQL & "   AND OrderCode   = '" & strOrdCd & "'    " & vbCrLf
                    SQL = SQL & "   And JStatus < '3'                       " & vbCrLf
                    
                    Call SetSQLData("결과저장", SQL, "A")
                    AdoCn.Execute SQL
                                        
                            
                End If
                RsLocal.MoveNext
            Loop
        End If
        
        RsLocal.Close
        
        SaveTransData_JWINFO = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_JWINFO = -1
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SaveTransData_JWINFO" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

Function SaveTransData_EASYS(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    Dim RsLocal         As ADODB.Recordset
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strPatNm        As String
    
    Dim strEqpcd        As String
    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strTestCdSub    As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strJudge        As String
    
On Error GoTo ErrHandle
    
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""

    With frmInterface
        SaveTransData_EASYS = -1
        
        strSaveSeq = Trim(GetText(SPD, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(SPD, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(SPD, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(SPD, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(SPD, argSpcRow, colPID))
        strPatNm = Trim(GetText(SPD, argSpcRow, colPNAME))
        strChartNo = Trim(GetText(SPD, argSpcRow, colCHARTNO))
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatNm) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMCODESUB,EQUIPRESULT,RESULT,REFJUDGE    " & vbCrLf
        SQL = SQL & "  FROM PATRESULT                                                               " & vbCrLf
        SQL = SQL & " WHERE EXAMDATE    = '" & strExamDate & "'                                     " & vbCrLf
        SQL = SQL & "   AND SAVESEQ     = " & strSaveSeq & vbCrLf
        SQL = SQL & "   AND BARCODE     = '" & strBarcode & "'                                      " & vbCrLf
        SQL = SQL & "   AND EXAMCODE    <> ''                                                       " & vbCrLf
        
        Set RsLocal = New ADODB.Recordset
        Set RsLocal = AdoCn_Local.Execute(SQL, , 1)
        If Not RsLocal.EOF = True And Not RsLocal.BOF = True Then
            Do Until RsLocal.EOF
                strEqpcd = RsLocal.Fields("EQUIPCODE").Value & ""
                'strOrdCd = RsLocal.Fields("ORDERCODE").Value & ""
                strTestCd = RsLocal.Fields("EXAMCODE").Value & ""
                'strTestCdSub = RsLocal.Fields("EXAMCODESUB").Value & ""
                sResult1 = RsLocal.Fields("EQUIPRESULT").Value & ""
                sResult2 = RsLocal.Fields("RESULT").Value & ""
                strJudge = RsLocal.Fields("REFJUDGE").Value & ""
                
                '-- 장비결과적용
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                    SQL = ""
                    SQL = SQL & "UPDATE H3LAB_RESULT    " & vbCrLf
                    SQL = SQL & "   SET STS_CD     = 'R'" & vbCrLf
                    If IsNumeric(sResult) Then
                        SQL = SQL & "      ,RESULT_VAL = '" & sResult & "'      " & vbCrLf '수치형결과
                    End If
                    SQL = SQL & "      ,RESULT_NM  = '" & sResult & "'      " & vbCrLf '(수치형 + 서술형 등 모든결과)
                    SQL = SQL & "      ,HL_GB      = '" & strJudge & "'     " & vbCrLf
                    SQL = SQL & " WHERE RECEPT_NO  = '" & strBarcode & "'   " & vbCrLf
                    SQL = SQL & "   And ORD_CD     = '" & strTestCd & "'    " & vbCrLf
                    'SQL = SQL & "   And STS_CD     = 'A'                    " & vbCrLf
                    
                    Call SetSQLData("결과저장", SQL, "A")
                    AdoCn.Execute SQL
                End If
                RsLocal.MoveNext
            Loop
        End If
        
        RsLocal.Close
        
        SaveTransData_EASYS = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_EASYS = -1
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SaveTransData_EASYS" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function


Function SaveTransData_MEDICHART(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    Dim RsLocal         As ADODB.Recordset
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strPatNm        As String
    
    Dim strEqpcd        As String
    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strTestCdSub    As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strJudge        As String
    Dim strYear         As String
    Dim strMonth        As String
    Dim strDay          As String
    
On Error GoTo ErrHandle
    
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""

    With frmInterface
        SaveTransData_MEDICHART = -1
        
        strSaveSeq = Trim(GetText(SPD, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(SPD, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(SPD, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(SPD, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(SPD, argSpcRow, colPID))
        strPatNm = Trim(GetText(SPD, argSpcRow, colPNAME))
        strChartNo = Trim(GetText(SPD, argSpcRow, colCHARTNO))
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatNm) = "" Then
            Exit Function
        End If
        
        strYear = Mid(strHospDate, 1, 4)
        strMonth = Mid(strHospDate, 5, 2)
        strDay = Mid(strHospDate, 7, 2)
        
        '-- Local에서 환자별로 결과값 가져오기
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMCODESUB,EQUIPRESULT,RESULT,REFJUDGE    " & vbCrLf
        SQL = SQL & "  FROM PATRESULT                                                               " & vbCrLf
        SQL = SQL & " WHERE EXAMDATE    = '" & strExamDate & "'                                     " & vbCrLf
        SQL = SQL & "   AND SAVESEQ     = " & strSaveSeq & vbCrLf
        SQL = SQL & "   AND BARCODE     = '" & strBarcode & "'                                      " & vbCrLf
        SQL = SQL & "   AND EXAMCODE    <> ''                                                       " & vbCrLf
        
        Set RsLocal = New ADODB.Recordset
        Set RsLocal = AdoCn_Local.Execute(SQL, , 1)
        If Not RsLocal.EOF = True And Not RsLocal.BOF = True Then
            Do Until RsLocal.EOF
                strEqpcd = RsLocal.Fields("EQUIPCODE").Value & ""
                strOrdCd = RsLocal.Fields("ORDERCODE").Value & ""
                strTestCd = RsLocal.Fields("EXAMCODE").Value & ""
                strTestCdSub = RsLocal.Fields("EXAMCODESUB").Value & ""
                sResult1 = RsLocal.Fields("EQUIPRESULT").Value & ""
                sResult2 = RsLocal.Fields("RESULT").Value & ""
                strJudge = RsLocal.Fields("REFJUDGE").Value & ""
                
                '-- 장비결과적용
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                  
                '-- 결과저장용 키 가져오기
                If strOrdCd = "" Then
                    strOrdCd = GetSampleSubITEM(strBarcode, strTestCd)
                End If
                
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" And strOrdCd <> "" Then
                    '-- 서버저장
                    SQL = ""
                    SQL = SQL & "Update TB_검사항목 "
                    SQL = SQL & "   Set 검사결과        = '" & sResult & "'                 " & vbCrLf
                    SQL = SQL & "     , 진료지원상태    = 5                                 " & vbCrLf '1 : 처치중, 5 : 완료
                    SQL = SQL & "     , 하이로우        = '" & strJudge & "'                " & vbCrLf
                    SQL = SQL & " Where 진료년          = '" & strYear & "'                 " & vbCrLf
                    SQL = SQL & "   and 진료월          = '" & strMonth & "'                " & vbCrLf
                    SQL = SQL & "   and 진료일          = '" & strDay & "'                  " & vbCrLf
                    SQL = SQL & "   and 챠트번호        = '" & strChartNo & "'              " & vbCrLf
                    SQL = SQL & "   And 처방코드        = '" & mGetP(strTestCd, 1, "|") & "'" & vbCrLf
                    SQL = SQL & "   And 서브코드        = '" & mGetP(strTestCd, 2, "|") & "'" & vbCrLf
                            
                    Call SetSQLData("결과저장", SQL, "A")
                    AdoCn.Execute SQL
                            
                End If
                RsLocal.MoveNext
            Loop
        End If
        
        RsLocal.Close
        
        SaveTransData_MEDICHART = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_MEDICHART = -1
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SaveTransData_MEDICHART" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

Function SaveTransData_UBCARE(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    Dim RsLocal         As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
    Dim strSex          As String
    Dim strAge          As String

    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strSubCode      As String
    Dim strEqpcd        As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strRefVal       As String
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    Dim strXMLBody      As String
    
    Dim strPName        As String
    Dim strPJumin       As String
    Dim strExamNo       As String
    Dim strSpcType      As String
    Dim strResult1      As String
    
On Error GoTo ErrHandle

    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""
    
    intRow = 0
    strResult1 = ""
    strXMLBody = ""
    blnSave = False
    
    SaveTransData_UBCARE = -1
    
    strSaveSeq = Trim(GetText(SPD, argSpcRow, colSAVESEQ))
    strExamDate = Trim(GetText(SPD, argSpcRow, colEXAMDATE))
    strExamDate = Format(strExamDate, "yyyymmdd")
    strHospDate = Trim(GetText(SPD, argSpcRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, argSpcRow, colBARCODE))
    strPatID = Trim(GetText(SPD, argSpcRow, colPID))
    strChartNo = Trim(GetText(SPD, argSpcRow, colCHARTNO))
    strIO = Trim(GetText(SPD, argSpcRow, colINOUT))
    
    If Trim(strBarcode) = "" Then
        Exit Function
    End If
    
    If Trim(strPatID) = "" Then
        Exit Function
    End If
    
    '-- Local에서 환자별로 결과값 가져오기
          SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT,EXAMNO,SAMPLETYPE,REFVALUE,PNAME,PJUMIN " & vbCrLf
    SQL = SQL & "  FROM UB_PATRESULT " & vbCrLf
    SQL = SQL & " WHERE EQUIPNO             = '" & gHOSP.MACHCD & "'            " & vbCrLf '장비코드
    'SQL = SQL & "   AND SAVESEQ            = " & strSaveSeq & vbCr                        '저장번호
    SQL = SQL & "   AND MID(EXAMDATE,1,8)   = '" & Mid(strExamDate, 1, 8) & "'  " & vbCrLf '검사일자
    SQL = SQL & "   AND BARCODE             = '" & strBarcode & "'              " & vbCrLf '바코드
    
    Call SetSQLData("로컬결과조회", SQL)
    
    Set RsLocal = New ADODB.Recordset
    Set RsLocal = AdoCn_Local.Execute(SQL, , 1)
    If Not RsLocal.EOF = True And Not RsLocal.BOF = True Then
        Do Until RsLocal.EOF
            strEqpcd = RsLocal.Fields("EQUIPCODE").Value & ""
            strOrdCd = RsLocal.Fields("ORDERCODE").Value & ""
            strTestCd = RsLocal.Fields("EXAMCODE").Value & ""
            sResult1 = RsLocal.Fields("EQUIPRESULT").Value & ""
            sResult2 = RsLocal.Fields("RESULT").Value & ""
            strExamNo = RsLocal.Fields("EXAMNO").Value & ""
            strSpcType = RsLocal.Fields("SAMPLETYPE").Value & ""
            strRefVal = RsLocal.Fields("REFVALUE").Value & ""
            strPName = RsLocal.Fields("PNAME").Value & ""
            strPJumin = RsLocal.Fields("PJUMIN").Value & ""
            
            If strIO = "입원" Then
                strIO = "1"
            ElseIf strIO = "외래" Then
                strIO = "0"
            End If
            
            '-- 결과적용
            If gHOSP.SAVELIS = "Y" Then
                sResult = sResult2
            Else
                sResult = sResult1
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                strXMLBody = strXMLBody & "<검사>"
                strXMLBody = strXMLBody & "<업체>" & "ACK" & "</업체>"
                strXMLBody = strXMLBody & "<요양기관번호>" & gHOSP.HOSPCD & "</요양기관번호>"
                strXMLBody = strXMLBody & "<차트번호>" & strChartNo & "</차트번호>"
                strXMLBody = strXMLBody & "<수진자명>" & strPName & "</수진자명>"
                strXMLBody = strXMLBody & "<주민등록번호>" & strPJumin & "</주민등록번호>"
                strXMLBody = strXMLBody & "<내원번호>" & strPatID & "</내원번호>"
                strXMLBody = strXMLBody & "<의뢰일>" & strHospDate & "</의뢰일>"
                strXMLBody = strXMLBody & "<검사번호>" & strExamNo & "</검사번호>"
                strXMLBody = strXMLBody & "<검사ID>" & strTestCd & "</검사ID>"
                strXMLBody = strXMLBody & "<업체검사ID>" & "" & "</업체검사ID>"
                strXMLBody = strXMLBody & "<검체>" & strSpcType & "</검체>"
            
                '소견형 결과(HBA1C, RBC-INDEX, WBC DIFF, 소변7종 등등... )
                If Mid(strTestCd, 1, 7) = "OPINION" Then
                    strXMLBody = strXMLBody & "<결과치>1</결과치>"
                    strXMLBody = strXMLBody & "<참조치>" & strRefVal & "</참조치>"
                    strXMLBody = strXMLBody & "<소견>" & sResult & "</소견>"
                Else
                    strXMLBody = strXMLBody & "<결과치>" & sResult & "</결과치>"
                    strXMLBody = strXMLBody & "<참조치>" & strRefVal & "</참조치>"
                    strXMLBody = strXMLBody & "<소견></소견>"
                End If
                strXMLBody = strXMLBody & "<결과일>" & Mid(strExamDate, 1, 8) & "</결과일>"
                strXMLBody = strXMLBody & "<입원외래구분>" & strIO & "</입원외래구분>"
                strXMLBody = strXMLBody & "</검사>"
            End If
            RsLocal.MoveNext
        Loop
    End If
    
    If strXMLBody <> "" Then
        Call SaveXMLFile_UBCARE(strXMLBody)
        strXMLBody = ""
        SaveTransData_UBCARE = 1
    End If
                
Exit Function

ErrHandle:
    SaveTransData_UBCARE = -1
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SaveTransData_UBCARE" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

Public Sub SaveXMLFile_UBCARE(strXMLBody As String, Optional argFlag As Integer = 0)
    Dim FilNum, FilNum1
    Dim FindFile    As String
    Dim TxtString1  As String
    Dim AllString1  As String
    Dim i As Long
    
    Dim strXmlLine      As String
    Dim strXmlAll       As String
    Dim strXmlAllBody   As String
    
    Dim strXml          As String
    Dim strXmlHeader    As String
    Dim strXmlTail      As String
    
    strXmlAll = ""
    strXmlAllBody = ""
    
    strXml = ""
    strXmlHeader = ""
    strXmlTail = ""
    
    strXmlHeader = "<?xml version=""1.0"" encoding=""euc-kr""?>" & vbCrLf & _
                   "<?xml-stylesheet type=""text/xsl"" href=C:\UBCare\SINAI\IF\Form\ExamIF_Form_05.xsl""?>" & vbCrLf & _
                   "<UBCare검사정보>"
    
    strXmlTail = "</UBCare검사정보>"
    
    If gHOSP.PARTCD = "C" Then
        'FindFile = Dir("C:\UBCare\SINAI\IF\ExamIF_Out1.xml")
        FindFile = Dir(gHOSP.XMLPATH & "\ExamIF_Out1.xml")
        
        If FindFile <> "" Then
            FilNum1 = FreeFile
            'Open "C:\UBCare\SINAI\IF\ExamIF_Out1.xml" For Input As FilNum1
            Open gHOSP.XMLPATH & "\ExamIF_Out1.xml" For Input As FilNum1
            
            Do While Not EOF(FilNum1)
                Input #FilNum1, strXmlLine
                strXmlAll = strXmlAll & strXmlLine
            Loop
    
            Close #FilNum1
            i = InStr(1, strXmlAll, "</UBCare검사정보>")
            strXmlAllBody = Mid(strXmlAll, 1, i - 1)
            strXml = strXmlAllBody & strXMLBody & strXmlTail
            
            'Kill "C:\UBCare\SINAI\IF\ExamIF_Out1.xml"
            Kill gHOSP.XMLPATH & "\ExamIF_Out1.xml"
        Else
            strXml = strXmlHeader & strXMLBody & strXmlTail
        End If
        
        FilNum = FreeFile
        
        If argFlag = 0 Then
            'Open "C:\UBCare\SINAI\IF\ExamIF_Out1.xml" For Output As FilNum
            Open gHOSP.XMLPATH & "\ExamIF_Out1.xml" For Output As FilNum
        Else
            'Open "C:\UBCare\SINAI\IF\ExamIF_Out1.xml" For Append As FilNum
            Open gHOSP.XMLPATH & "\ExamIF_Out1.xml" For Append As FilNum
        End If
    ElseIf gHOSP.PARTCD = "H" Then
        'FindFile = Dir("C:\UBCare\SINAI\IF\ExamIF_Out3.xml")
        FindFile = Dir(gHOSP.XMLPATH & "\ExamIF_Out3.xml")
        
        If FindFile <> "" Then
            FilNum1 = FreeFile
        '    Open "C:\UBCare\SINAI\IF\ExamIF_Out3.xml" For Input As FilNum1
            Open gHOSP.XMLPATH & "\ExamIF_Out3.xml" For Input As FilNum1
            
            Do While Not EOF(FilNum1)
                Input #FilNum1, strXmlLine
                strXmlAll = strXmlAll & strXmlLine
            Loop
    
            Close #FilNum1
            i = InStr(1, strXmlAll, "</UBCare검사정보>")
            strXmlAllBody = Mid(strXmlAll, 1, i - 1)
            strXml = strXmlAllBody & strXMLBody & strXmlTail
         '   Kill "C:\UBCare\SINAI\IF\ExamIF_Out3.xml"
            Kill gHOSP.XMLPATH & "\ExamIF_Out3.xml"
        Else
            strXml = strXmlHeader & strXMLBody & strXmlTail
        End If
        
        FilNum = FreeFile
        
        If argFlag = 0 Then
          '  Open "C:\UBCare\SINAI\IF\ExamIF_Out3.xml" For Output As FilNum
            Open gHOSP.XMLPATH & "\ExamIF_Out3.xml" For Output As FilNum
        Else
           ' Open "C:\UBCare\SINAI\IF\ExamIF_Out3.xml" For Append As FilNum
            Open gHOSP.XMLPATH & "\ExamIF_Out3.xml" For Append As FilNum
        End If

    Else
       ' FindFile = Dir("C:\UBCare\SINAI\IF\ExamIF_Out2.xml")
        FindFile = Dir(gHOSP.XMLPATH & "\ExamIF_Out2.xml")
        
        If FindFile <> "" Then
            FilNum1 = FreeFile
       '     Open "C:\UBCare\SINAI\IF\ExamIF_Out2.xml" For Input As FilNum1
            Open gHOSP.XMLPATH & "\ExamIF_Out2.xml" For Input As FilNum1
            
            Do While Not EOF(FilNum1)
                Input #FilNum1, strXmlLine
                strXmlAll = strXmlAll & strXmlLine
            Loop
    
            Close #FilNum1
            i = InStr(1, strXmlAll, "</UBCare검사정보>")
            strXmlAllBody = Mid(strXmlAll, 1, i - 1)
            strXml = strXmlAllBody & strXMLBody & strXmlTail
        '    Kill "C:\UBCare\SINAI\IF\ExamIF_Out2.xml"
            Kill gHOSP.XMLPATH & "\ExamIF_Out2.xml"
        Else
            strXml = strXmlHeader & strXMLBody & strXmlTail
        End If
        
        FilNum = FreeFile
        
        If argFlag = 0 Then
        '    Open "C:\UBCare\SINAI\IF\ExamIF_Out2.xml" For Output As FilNum
            Open gHOSP.XMLPATH & "\ExamIF_Out2.xml" For Output As FilNum
        Else
        '    Open "C:\UBCare\SINAI\IF\ExamIF_Out2.xml" For Append As FilNum
            Open gHOSP.XMLPATH & "\ExamIF_Out2.xml" For Append As FilNum
        End If

    End If
    
    Print #FilNum, strXml
    Close FilNum
    
    Call SetSQLData("결과저장", strXml, "A")
    
    
End Sub

Function SaveTransData_BITJSON(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    Dim RsLocal         As ADODB.Recordset
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strPatNm        As String
    
    Dim strEqpcd        As String
    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strTestCdSub    As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strJudge        As String
    Dim strYear         As String
    Dim strMonth        As String
    Dim strDay          As String
    
    Dim strParam(7) As Variant
    Dim strReturn   As Variant
    
    
On Error GoTo ErrHandle
    
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""

    With frmInterface
        SaveTransData_BITJSON = -1
        
        strSaveSeq = Trim(GetText(SPD, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(SPD, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(SPD, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(SPD, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(SPD, argSpcRow, colPID))
        strPatNm = Trim(GetText(SPD, argSpcRow, colPNAME))
        strChartNo = Trim(GetText(SPD, argSpcRow, colCHARTNO))
        
        'If Trim(strChartNo) = "" Then
        '    Exit Function
        'End If
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        strYear = Mid(strHospDate, 1, 4)
        strMonth = Mid(strHospDate, 5, 2)
        strDay = Mid(strHospDate, 7, 2)
        
        '-- Local에서 환자별로 결과값 가져오기
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMCODESUB,EQUIPRESULT,RESULT,REFJUDGE    " & vbCrLf
        SQL = SQL & "  FROM PATRESULT                                                               " & vbCrLf
        SQL = SQL & " WHERE EXAMDATE    = '" & strExamDate & "'                                     " & vbCrLf
        SQL = SQL & "   AND SAVESEQ     = " & strSaveSeq & vbCrLf
        SQL = SQL & "   AND BARCODE     = '" & strBarcode & "'                                      " & vbCrLf
        SQL = SQL & "   AND EXAMCODE    <> ''                                                       " & vbCrLf
        
        Set RsLocal = New ADODB.Recordset
        Set RsLocal = AdoCn_Local.Execute(SQL, , 1)
        If Not RsLocal.EOF = True And Not RsLocal.BOF = True Then
            Do Until RsLocal.EOF
                strEqpcd = RsLocal.Fields("EQUIPCODE").Value & ""
                strOrdCd = RsLocal.Fields("ORDERCODE").Value & ""
                strTestCd = RsLocal.Fields("EXAMCODE").Value & ""
                strTestCdSub = RsLocal.Fields("EXAMCODESUB").Value & ""
                sResult1 = RsLocal.Fields("EQUIPRESULT").Value & ""
                sResult2 = RsLocal.Fields("RESULT").Value & ""
                strJudge = RsLocal.Fields("REFJUDGE").Value & ""
                
                '-- 장비결과적용
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                  
                
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                    '-- 서버저장
'"brcdLablNo : 검체번호
'exmnCd:  검사코드
'realRslt:  실제결과
'viewRslt:  검사결과
'eqpmCd:  장비코드
'eqpmFlag:  장비FLAG
'examDt:  검사일시
'exmnId:  검사자
'{
'  brcdLablNo : ""1820027311""
'  exmnCd : ""L3011""
'  realRslt :""100""
'  viewRslt : ""100""
'  eqpmCd : ""011""
'  eqpmFlag : ""1""
'  examDt : ""20180504010101""
'  exmnId : ""test""
'}"

                    strParam(0) = strBarcode
                    strParam(1) = strTestCd
                    strParam(2) = sResult1
                    strParam(3) = sResult2
                    strParam(4) = gHOSP.MACHCD
                    strParam(5) = ""
                    strParam(6) = Format(Now, "yyyymmdd")
                    strParam(7) = gHOSP.USERID
                    Call SetSQLData("결과저장", strParam(0) & "," & strParam(1) & "," & strParam(2) & "," & strParam(3) & "," & strParam(4) & "," & strParam(5) & "," & strParam(6) & "," & strParam(7), "A")
                    strReturn = JsonSend("PATSAVE", strParam)
                    Call SetSQLData("결과저장", strReturn, "A")
                    AdoCn.Execute SQL
                            
                End If
                RsLocal.MoveNext
            Loop
        End If
        
        RsLocal.Close
        
        SaveTransData_BITJSON = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_BITJSON = -1
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SaveTransData_BITJSON" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'Function SaveTransData_HEALTH(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
'    Dim oSOAP       As MSSOAPLib30.SoapClient30
'    Dim strHeader   As String
'    Dim strBody     As String
'    Dim strParam    As String
'    Dim strReturn   As Variant
'    Dim strHData    As Variant
'
'    Dim RsLocal         As ADODB.Recordset
'
'    Dim strSaveSeq      As String
'    Dim strExamDate     As String
'    Dim strHospDate     As String
'    Dim strBarcode      As String
'    Dim strChartNo      As String
'    Dim strPatID        As String
'    Dim strPatNm        As String
'
'    Dim strEqpcd        As String
'    Dim strOrdCd        As String
'    Dim strTestCd       As String
'    Dim strTestCdSub    As String
'    Dim sResult         As String
'    Dim sResult1        As String
'    Dim sResult2        As String
'    Dim strJudge        As String
'    Dim strYear         As String
'    Dim strMonth        As String
'    Dim strDay          As String
'
''    Dim strParam(7)     As Variant
''    Dim strReturn       As Variant
'    Dim j               As Integer
'
'On Error GoTo ErrHandle
'
'    strJudge = ""
'    sResult = ""
'    sResult1 = ""
'    sResult2 = ""
'
'    strParam = ""
'    strHeader = ""
'    strBody = ""
'    j = 0
'
'    With frmInterface
'        SaveTransData_HEALTH = -1
'
'        strSaveSeq = Trim(GetText(SPD, argSpcRow, colSAVESEQ))
'        strExamDate = Trim(GetText(SPD, argSpcRow, colEXAMDATE))
'        strHospDate = Trim(GetText(SPD, argSpcRow, colHOSPDATE))
'        strBarcode = Trim(GetText(SPD, argSpcRow, colBARCODE))
'        strPatID = Trim(GetText(SPD, argSpcRow, colPID))
'        strPatNm = Trim(GetText(SPD, argSpcRow, colPNAME))
'        strChartNo = Trim(GetText(SPD, argSpcRow, colCHARTNO))
'
'        If Trim(strBarcode) = "" Then
'            Exit Function
'        End If
'
'        '-- Local에서 환자별로 결과값 가져오기
'              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMCODESUB,EQUIPRESULT,RESULT,REFJUDGE    " & vbCrLf
'        SQL = SQL & "  FROM PATRESULT                                                               " & vbCrLf
'        SQL = SQL & " WHERE EXAMDATE    = '" & strExamDate & "'                                     " & vbCrLf
'        SQL = SQL & "   AND SAVESEQ     = " & strSaveSeq & vbCrLf
'        SQL = SQL & "   AND BARCODE     = '" & strBarcode & "'                                      " & vbCrLf
'        SQL = SQL & "   AND EXAMCODE    <> ''                                                       " & vbCrLf
'
'        Set RsLocal = New ADODB.Recordset
'        Set RsLocal = AdoCn_Local.Execute(SQL, , 1)
'        If Not RsLocal.EOF = True And Not RsLocal.BOF = True Then
'            Do Until RsLocal.EOF
'                strEqpcd = RsLocal.Fields("EQUIPCODE").Value & ""
'                strOrdCd = RsLocal.Fields("ORDERCODE").Value & ""
'                strTestCd = RsLocal.Fields("EXAMCODE").Value & ""
'                strTestCdSub = RsLocal.Fields("EXAMCODESUB").Value & ""
'                sResult1 = RsLocal.Fields("EQUIPRESULT").Value & ""
'                sResult2 = RsLocal.Fields("RESULT").Value & ""
'                strJudge = RsLocal.Fields("REFJUDGE").Value & ""
'
'                '-- 장비결과적용
'                If gHOSP.SAVELIS = "Y" Then
'                    sResult = sResult2
'                Else
'                    sResult = sResult1
'                End If
'
'
'                If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
'                    j = j + 1
'                    strBody = strBody & "OBX|" & CStr(j) & "|" & "ST" & "|" & strTestCd & "||" & sResult & "||||||R" & Chr(13)
'                End If
'                RsLocal.MoveNext
'            Loop
'        End If
'
'        RsLocal.Close
'
'        If strParam <> "" Then
'            strHeader = ""
'            strHeader = strHeader & "MSH|^~\&|HL7|MMS|||" & Format(Now, "yyyymmddhhmmss") & "||ORU^R01|1a082e2:10e59b48c04:-2cf9:27695009|P|2.3||||||8859/1 " & Chr(13)
'            strHeader = strHeader & "PID|||" & strBarcode & "^" & gHOSP.MACHCD & "^" & gHOSP.USERID & "^^^DefaultDomain^PI" & Chr(13)
'            strHeader = strHeader & "PV1||E|" & gHOSP.HOSPCD & Chr(13)
'            strHeader = strHeader & "OBR|1||||||" & Format(Now, "yyyymmddhhmmss") & Chr(13)
'
'            strParam = Chr(11) & strHeader & strBody
'
'            Set oSOAP = New MSSOAPLib30.SoapClient30
'            oSOAP.ClientProperty("ServerHTTPRequest") = True
'            oSOAP.MSSoapInit gHEALTH.INITURL
'
'            Call SetSQLData("결과저장", strParam, "A")
'            strParam = MakeB64(strParam)
'            strReturn = oSOAP.UpdateRst(strParam)
'            strReturn = MakeUB64(strReturn)
'            Call SetSQLData("결과저장", strReturn, "A")
'            Set oSOAP = Nothing
'            DoEvents
'        End If
'
'        SaveTransData_HEALTH = 1
'
'    End With
'
'Exit Function
'
'ErrHandle:
'    SaveTransData_HEALTH = -1
'    Screen.MousePointer = 0
'
'    strErrMsg = ""
'    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SaveTransData_HEALTH" & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
'    frmErrMsg.txtErr = vbNewLine & strErrMsg
'    frmErrMsg.Show
'
'End Function

Function SaveTransData_NEOSENSE(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    Dim RsLocal         As ADODB.Recordset
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strPatNm        As String
    
    Dim strEqpcd        As String
    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strTestCdSub    As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strJudge        As String
    Dim strYear         As String
    Dim strMonth        As String
    Dim strDay          As String
    
On Error GoTo ErrHandle
    
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""

    With frmInterface
        SaveTransData_NEOSENSE = -1
        
        strSaveSeq = Trim(GetText(SPD, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(SPD, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(SPD, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(SPD, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(SPD, argSpcRow, colPID))
        strPatNm = Trim(GetText(SPD, argSpcRow, colPNAME))
        strChartNo = Trim(GetText(SPD, argSpcRow, colCHARTNO))
        
        If Trim(strChartNo) = "" Then
            Exit Function
        End If
        
        If Trim(strPatNm) = "" Then
            Exit Function
        End If
        
        strYear = Mid(strHospDate, 1, 4)
        strMonth = Mid(strHospDate, 5, 2)
        strDay = Mid(strHospDate, 7, 2)
        
        '-- Local에서 환자별로 결과값 가져오기
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMCODESUB,EQUIPRESULT,RESULT,REFJUDGE    " & vbCrLf
        SQL = SQL & "  FROM PATRESULT                                                               " & vbCrLf
        SQL = SQL & " WHERE EXAMDATE    = '" & strExamDate & "'                                     " & vbCrLf
        SQL = SQL & "   AND SAVESEQ     = " & strSaveSeq & vbCrLf
        SQL = SQL & "   AND BARCODE     = '" & strBarcode & "'                                      " & vbCrLf
        SQL = SQL & "   AND EXAMCODE    <> ''                                                       " & vbCrLf
        
        Set RsLocal = New ADODB.Recordset
        Set RsLocal = AdoCn_Local.Execute(SQL, , 1)
        If Not RsLocal.EOF = True And Not RsLocal.BOF = True Then
            Do Until RsLocal.EOF
                strEqpcd = RsLocal.Fields("EQUIPCODE").Value & ""
                strOrdCd = RsLocal.Fields("ORDERCODE").Value & ""
                strTestCd = RsLocal.Fields("EXAMCODE").Value & ""
                strTestCdSub = RsLocal.Fields("EXAMCODESUB").Value & ""
                sResult1 = RsLocal.Fields("EQUIPRESULT").Value & ""
                sResult2 = RsLocal.Fields("RESULT").Value & ""
                strJudge = RsLocal.Fields("REFJUDGE").Value & ""
                
                '-- 장비결과적용
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                  
                '-- 결과저장용 키 가져오기
                If strOrdCd = "" Then
                    strOrdCd = GetSampleSubITEM(strBarcode, strTestCd)
                End If
                
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                    '-- 서버저장
                    If strJudge = "L" Then
                        strJudge = "2"
                    ElseIf strJudge = "H" Then
                        strJudge = "1"
                    Else
                        strJudge = "0"
                    End If
                    
                    SQL = ""
                    SQL = SQL & "Update TB_검사항목 "
                    SQL = SQL & "   Set 검사결과        = '" & sResult & "'                 " & vbCrLf
'                    SQL = SQL & "     , 진료지원상태    = 5                                 " & vbCrLf '1 : 처치중, 5 : 완료
                    SQL = SQL & "     , 하이로우        = '" & strJudge & "'                " & vbCrLf
                    SQL = SQL & " Where 진료년          = '" & strYear & "'                 " & vbCrLf
                    SQL = SQL & "   and 진료월          = '" & strMonth & "'                " & vbCrLf
                    SQL = SQL & "   and 진료일          = '" & strDay & "'                  " & vbCrLf
                    SQL = SQL & "   and 챠트번호        = '" & strChartNo & "'              " & vbCrLf
                    SQL = SQL & "   And 처방코드        = '" & mGetP(strTestCd, 1, "|") & "'" & vbCrLf
                    SQL = SQL & "   And 서브코드        = '" & mGetP(strTestCd, 2, "|") & "'" & vbCrLf
                            
                    Call SetSQLData("결과저장", SQL, "A")
                    AdoCn.Execute SQL
                            
                End If
                RsLocal.MoveNext
            Loop
        End If
        
        RsLocal.Close
        
        SaveTransData_NEOSENSE = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_NEOSENSE = -1
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SaveTransData_NEOSENSE" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

Function SaveTransData_LABSPEAR(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    Dim RsLocal         As ADODB.Recordset
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strPatNm        As String
    
    Dim strEqpcd        As String
    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strTestCdSub    As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strJudge        As String
    Dim strCmnt         As String
    Dim strRet          As String
    
    Dim Prm1            As New ADODB.Parameter
    Dim Prm2            As New ADODB.Parameter
    Dim Prm3            As New ADODB.Parameter
    Dim prm4            As New ADODB.Parameter
    Dim prm5            As New ADODB.Parameter
    Dim prm6            As New ADODB.Parameter
    Dim prm7            As New ADODB.Parameter
    Dim prm8            As New ADODB.Parameter
    Dim prm9            As New ADODB.Parameter
    Dim prm10           As New ADODB.Parameter
    Dim prm11           As New ADODB.Parameter
    Dim prm12           As New ADODB.Parameter
    Dim prm13           As New ADODB.Parameter
    Dim prm14           As New ADODB.Parameter
    Dim prm15           As New ADODB.Parameter
    Dim prm16           As New ADODB.Parameter
    
    Dim prmcmt1         As New ADODB.Parameter
    Dim prmcmt2         As New ADODB.Parameter
    Dim prmcmt3         As New ADODB.Parameter
    Dim prmcmt4         As New ADODB.Parameter
    Dim prmcmt5         As New ADODB.Parameter
    Dim prmcmt6         As New ADODB.Parameter
    Dim prmcmt7         As New ADODB.Parameter
    Dim prmcmt8         As New ADODB.Parameter
    Dim prmcmt9         As New ADODB.Parameter
    
On Error GoTo ErrHandle
    
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""
    strCmnt = ""
    
    With frmInterface
        SaveTransData_LABSPEAR = -1
        
        strSaveSeq = Trim(GetText(SPD, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(SPD, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(SPD, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(SPD, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(SPD, argSpcRow, colPID))
        strPatNm = Trim(GetText(SPD, argSpcRow, colPNAME))
        strChartNo = Trim(GetText(SPD, argSpcRow, colCHARTNO))
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatNm) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMCODESUB,EQUIPRESULT,RESULT,REFJUDGE    " & vbCrLf
        SQL = SQL & "  FROM PATRESULT                                                               " & vbCrLf
        SQL = SQL & " WHERE EXAMDATE    = '" & strExamDate & "'                                     " & vbCrLf
        SQL = SQL & "   AND SAVESEQ     = " & strSaveSeq & vbCrLf
        SQL = SQL & "   AND BARCODE     = '" & strBarcode & "'                                      " & vbCrLf
        SQL = SQL & "   AND EXAMCODE    <> ''                                                       " & vbCrLf
        
        Set RsLocal = New ADODB.Recordset
        Set RsLocal = AdoCn_Local.Execute(SQL, , 1)
        If Not RsLocal.EOF = True And Not RsLocal.BOF = True Then
            Do Until RsLocal.EOF
                strEqpcd = RsLocal.Fields("EQUIPCODE").Value & ""
                strOrdCd = RsLocal.Fields("ORDERCODE").Value & ""
                strTestCd = RsLocal.Fields("EXAMCODE").Value & ""
                strTestCdSub = RsLocal.Fields("EXAMCODESUB").Value & ""
                sResult1 = RsLocal.Fields("EQUIPRESULT").Value & ""
                sResult2 = RsLocal.Fields("RESULT").Value & ""
                strJudge = RsLocal.Fields("REFJUDGE").Value & ""
                
                '-- 장비결과적용
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                
                If strPatID <> "" And strTestCd <> "" And sResult <> "" Then
                    '-- 서버저장
                    Set AdoCmd = New ADODB.Command
                    Set AdoCmd.ActiveConnection = AdoCn
                    With AdoCmd
                        .CommandTimeout = 15
                        .CommandText = "sp_검사값저장"
                        .CommandType = adCmdStoredProc
                        
                        Set Prm1 = .CreateParameter("receiptdate", adDate, adParamInput, 30, Format(strHospDate, "####-##-##"))
                        .Parameters.Append Prm1
                        
                        Set Prm2 = .CreateParameter("receiptnum", adVarChar, adParamInput, 30, strPatID)
                        .Parameters.Append Prm2
                        
                        Set Prm3 = .CreateParameter("labcode", adVarChar, adParamInput, 30, strTestCd)
                        .Parameters.Append Prm3
                        
                        Set prm4 = .CreateParameter("resultvalue", adVarChar, adParamInput, 4000, sResult)
                        .Parameters.Append prm4
                        
                        Set prm5 = .CreateParameter("resultvalue2", adVarChar, adParamInput, 50, "")
                        .Parameters.Append prm5
                        
                        Set prm6 = .CreateParameter("resultvalue3", adVarChar, adParamInput, 50, "")
                        .Parameters.Append prm6
                        
                        Set prm7 = .CreateParameter("abnormal", adVarChar, adParamInput, 30, strJudge)
                        .Parameters.Append prm7
                        
                        Set prm8 = .CreateParameter("panic", adVarChar, adParamInput, 30, "")
                        .Parameters.Append prm8
                        
                        Set prm9 = .CreateParameter("critical", adVarChar, adParamInput, 30, "")
                        .Parameters.Append prm9
                        
                        Set prm10 = .CreateParameter("amr", adVarChar, adParamInput, 30, "")
                        .Parameters.Append prm10
                        
                        Set prm11 = .CreateParameter("crr", adVarChar, adParamInput, 30, "")
                        .Parameters.Append prm11
                        
                        Set prm12 = .CreateParameter("machinecode", adVarChar, adParamInput, 30, gHOSP.MACHCD)
                        .Parameters.Append prm12
                        
                        Set prm13 = .CreateParameter("employeecode", adVarChar, adParamInput, 30, gHOSP.USERID)
                        .Parameters.Append prm13
                        
                        Set prm14 = .CreateParameter("computerid", adVarChar, adParamInput, 30, CStr(frmInterface.wSck.LocalIP))
                        .Parameters.Append prm14
                        
                        Set prm15 = .CreateParameter("overwrite", adVarChar, adParamInput, 1, "")
                        .Parameters.Append prm15
                        
                        Set prm16 = .CreateParameter("updatecount", adVarChar, adParamInputOutput, 100, 0)
                        .Parameters.Append prm16
    
                        .Execute
                        
                        strRet = .Parameters("updatecount").Value
                        If strRet = "0" Or strRet = "1" Then
                            '-- 저장성공
                            frmInterface.lblIFStatus.Caption = strPatID & " 검사결과 저장"
                            Set AdoCmd = Nothing
                            SaveTransData_LABSPEAR = 1
                        Else
                            '-- 저장실패
                            frmInterface.lblIFStatus.Caption = strPatID & " 검사결과 저장오류"
                            Set AdoCmd = Nothing
                            SaveTransData_LABSPEAR = -1
                        End If
                        
                        SQL = ""
                        SQL = SQL & Format(strHospDate, "####-##-##") & "," & vbCrLf
                        SQL = SQL & strPatID & "," & vbCrLf
                        SQL = SQL & strTestCd & "," & vbCrLf
                        SQL = SQL & sResult & "," & vbCrLf
                        SQL = SQL & "" & "," & vbCrLf
                        SQL = SQL & strJudge & "," & vbCrLf
                        SQL = SQL & "" & "," & vbCrLf
                        SQL = SQL & "" & "," & vbCrLf
                        SQL = SQL & "" & "," & vbCrLf
                        SQL = SQL & "" & "," & vbCrLf
                        SQL = SQL & gHOSP.MACHCD & "," & vbCrLf
                        SQL = SQL & gHOSP.USERID & "," & vbCrLf
                        SQL = SQL & CStr(frmInterface.wSck.LocalIP) & "," & vbCrLf
                        SQL = SQL & "" & "," & vbCrLf
                        SQL = SQL & "0" & "," & vbCrLf
                        SQL = SQL & "strRet:" & strRet & vbCrLf
                
                        Call SetSQLData("결과저장", SQL, "A")
                    End With
                End If
                RsLocal.MoveNext
            Loop
        End If
        
        '-- 학부메모저장
        If strPatID <> "" And strCmnt <> "" And (strRet = "0" Or strRet = "1") Then
            Set AdoCmd = New ADODB.Command
            Set AdoCmd.ActiveConnection = AdoCn
            With AdoCmd
                .CommandTimeout = 15
                .CommandText = "sp_검사학부메모저장"
                .CommandType = adCmdStoredProc
                
                Set prmcmt1 = .CreateParameter("receiptdate", adDate, adParamInput, 30, Format(strHospDate, "####-##-##"))
                .Parameters.Append prmcmt1
                
                Set prmcmt2 = .CreateParameter("receiptnum", adVarChar, adParamInput, 30, strPatID)
                .Parameters.Append prmcmt2
                
                Set prmcmt3 = .CreateParameter("labDeptcode", adVarChar, adParamInput, 30, gHOSP.PARTCD) 'C2
                .Parameters.Append prmcmt3
                
                Set prmcmt4 = .CreateParameter("labmemo", adVarChar, adParamInput, 2000, strCmnt)
                .Parameters.Append prmcmt4
                
                Set prmcmt5 = .CreateParameter("employeecode", adVarChar, adParamInput, 30, gHOSP.USERID)
                .Parameters.Append prmcmt5
                
                Set prmcmt6 = .CreateParameter("computerid", adVarChar, adParamInput, 30, CStr(frmInterface.wSck.LocalIP))
                .Parameters.Append prmcmt6
                
                Set prmcmt7 = .CreateParameter("overwrite", adVarChar, adParamInput, 1, "")
                .Parameters.Append prmcmt7
                
                Set prmcmt8 = .CreateParameter("updatecount", adVarChar, adParamInputOutput, 100, 0)
                .Parameters.Append prmcmt8
                
                .Execute
                
                strRet = .Parameters("updatecount").Value
                If strRet = "0" Or strRet = "1" Then
                    '-- 저장성공
                    Set AdoCmd = Nothing
                    SaveTransData_LABSPEAR = 1
                    
                Else
                    '-- 저장실패
                    'MsgBox "검사결과 저장오류 " & .Parameters("updatecount").Value, vbInformation + vbOKOnly
                    Set AdoCmd = Nothing
                    SaveTransData_LABSPEAR = -1
                End If
            End With
        End If
        RsLocal.Close
        
        SaveTransData_LABSPEAR = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_LABSPEAR = -1
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SaveTransData_LABSPEAR" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

Function SaveTransData_SCL(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    Dim RsLocal         As ADODB.Recordset
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strPatNm        As String
    
    Dim strEqpcd        As String
    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strTestCdSub    As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strJudge        As String
    Dim strCmnt         As String
    Dim strRet          As String
    
    Dim Prm1            As New ADODB.Parameter
    Dim Prm2            As New ADODB.Parameter
    Dim Prm3            As New ADODB.Parameter

On Error GoTo ErrHandle
    
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""
    strCmnt = ""
    
    With frmInterface
        SaveTransData_SCL = -1
        
        strSaveSeq = Trim(GetText(SPD, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(SPD, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(SPD, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(SPD, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(SPD, argSpcRow, colPID))
        strPatNm = Trim(GetText(SPD, argSpcRow, colPNAME))
        strChartNo = Trim(GetText(SPD, argSpcRow, colCHARTNO))
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatNm) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMCODESUB,EQUIPRESULT,RESULT,REFJUDGE    " & vbCrLf
        SQL = SQL & "  FROM PATRESULT                                                               " & vbCrLf
        SQL = SQL & " WHERE EXAMDATE    = '" & strExamDate & "'                                     " & vbCrLf
        SQL = SQL & "   AND SAVESEQ     = " & strSaveSeq & vbCrLf
        SQL = SQL & "   AND BARCODE     = '" & strBarcode & "'                                      " & vbCrLf
        SQL = SQL & "   AND EXAMCODE    <> ''                                                       " & vbCrLf
        
        Set RsLocal = New ADODB.Recordset
        Set RsLocal = AdoCn_Local.Execute(SQL, , 1)
        If Not RsLocal.EOF = True And Not RsLocal.BOF = True Then
            Do Until RsLocal.EOF
                strEqpcd = RsLocal.Fields("EQUIPCODE").Value & ""
                strOrdCd = RsLocal.Fields("ORDERCODE").Value & ""
                strTestCd = RsLocal.Fields("EXAMCODE").Value & ""
                strTestCdSub = RsLocal.Fields("EXAMCODESUB").Value & ""
                sResult1 = RsLocal.Fields("EQUIPRESULT").Value & ""
                sResult2 = RsLocal.Fields("RESULT").Value & ""
                strJudge = RsLocal.Fields("REFJUDGE").Value & ""
                
                '-- 장비결과적용
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                
                If strPatID <> "" And strTestCd <> "" And sResult <> "" Then
                    '-- 서버저장
                    SQL = ""
                    SQL = SQL & "Update LisiLib.Minterface                      " & vbCrLf
                    SQL = SQL & "   Set Result      = '" & Trim(sResult) & "'   " & vbCrLf
                    SQL = SQL & "     , Rltflag     = 'N'                       " & vbCrLf
                    SQL = SQL & "     , Updtdate    = (select substring(char(curdate()),1,4) || substring(char(curdate()),6,2) || substring(char(curdate()),9,2) || substring(char(curtime()),4,2) || substring(char(curtime()),7,2) || substring(char(curtime()),10,2) from sysibm.sysdummy1) " & vbCrLf
                    SQL = SQL & "     , Testercode  = '" & gHOSP.USERID & "'    " & vbCrLf
                    SQL = SQL & "     , Flag        = '2'                       " & vbCrLf
                    SQL = SQL & "     , Updtempl    = '" & gHOSP.USERID & "'    " & vbCrLf
                    '코멘트
                    If mResult.CMNTCD <> "" Then
                        SQL = SQL & "     , frltcode    = '" & mResult.CMNTCD & "'" & vbCrLf
                        mResult.CMNTCD = ""
                    End If
                    SQL = SQL & " Where barcodeno   = '" & strBarcode & "'      " & vbCrLf
                    SQL = SQL & "   And mcode       = '" & gHOSP.MACHCD & "'    " & vbCrLf
                    SQL = SQL & "   And itemcode    = '" & Mid(strTestCd, 1, 5) & "'" & vbCrLf
                    If Len(strTestCd) > 5 Then
                       SQL = SQL & "   And dcode = '" & Mid(strTestCd, 6) & "'"
                    End If
                    
                    
                    '코멘트 : frltcode = 코드
                    '판정   :
                    
                    Call SetSQLData("결과저장", SQL, "A")
                    AdoCn.Execute SQL
                    
                End If
                RsLocal.MoveNext
            Loop
        End If
        
        '-- 상태저장 (결과 저장이 완료되면 해당 procedure를 call 한다)
        'batch slrtrm55p(pmach : char(3) => 장비코드,
        '                perr  : char(1) => 인증확인 및 에러코드),
        '
        'real  slrtrm56p(pbarc : char(12) => 바코드번호,
        '                pmach : char(3) => 장비코드,
        '                perr  : char(1) => 인증확인 및 에러코드)
        
        If strBarcode <> "" Then 'And (strRet = "0" Or strRet = "1") Then
            Set AdoCmd = New ADODB.Command
            Set AdoCmd.ActiveConnection = AdoCn
            With AdoCmd
                .CommandTimeout = 15
                .CommandText = "SLRTRM56P"
                .CommandType = adCmdStoredProc
                
                Set Prm1 = .CreateParameter("pbarc", adChar, adParamInput, 12, strBarcode)
                .Parameters.Append Prm1
                
                'Set Prm2 = .CreateParameter("pmach", adChar, adParamInput, 3, gHOSP.MACHCD)
                Set Prm2 = .CreateParameter("pmach", adChar, adParamInput, 3, mResult.EqpCd)
                .Parameters.Append Prm2
                
                Set Prm3 = .CreateParameter("perr", adChar, adParamOutput, 1, "")
                .Parameters.Append Prm3
                
                .Execute
                
                Set AdoCmd = Nothing
                    
                Call SetSQLData("결과저장", "프로시져호출", "A")
            
            End With
        End If
        RsLocal.Close
        
        SaveTransData_SCL = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_SCL = -1
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SaveTransData_SCL" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function


Function SaveTransData_BIT70(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    Dim RsLocal         As ADODB.Recordset
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strPatNm        As String
    
    Dim strEqpcd        As String
    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strTestCdSub    As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strJudge        As String
    
    Dim strDate         As String
    Dim strTime         As String
    Dim blnSave         As Boolean
    
On Error GoTo ErrHandle
    
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""
    blnSave = False
    
    With frmInterface
        SaveTransData_BIT70 = -1
        
        strSaveSeq = Trim(GetText(SPD, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(SPD, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(SPD, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(SPD, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(SPD, argSpcRow, colPID))
        strPatNm = Trim(GetText(SPD, argSpcRow, colPNAME))
        strChartNo = Trim(GetText(SPD, argSpcRow, colCHARTNO))
        
        If Trim(strPatID) = "" Then
            Exit Function
        End If
        
        If Trim(strChartNo) = "" Then
            Exit Function
        End If
        
        If Trim(strPatNm) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMCODESUB,EQUIPRESULT,RESULT,REFJUDGE    " & vbCrLf
        SQL = SQL & "  FROM PATRESULT                                                               " & vbCrLf
        SQL = SQL & " WHERE EXAMDATE    = '" & strExamDate & "'                                     " & vbCrLf
        SQL = SQL & "   AND SAVESEQ     = " & strSaveSeq & vbCrLf
        SQL = SQL & "   AND BARCODE     = '" & strBarcode & "'                                      " & vbCrLf
        SQL = SQL & "   AND EXAMCODE    <> ''                                                       " & vbCrLf
        
        Set RsLocal = New ADODB.Recordset
        Set RsLocal = AdoCn_Local.Execute(SQL, , 1)
        If Not RsLocal.EOF = True And Not RsLocal.BOF = True Then
            Do Until RsLocal.EOF
                strEqpcd = RsLocal.Fields("EQUIPCODE").Value & ""
                strOrdCd = RsLocal.Fields("ORDERCODE").Value & ""
                strTestCd = RsLocal.Fields("EXAMCODE").Value & ""
                strTestCdSub = RsLocal.Fields("EXAMCODESUB").Value & ""
                sResult1 = RsLocal.Fields("EQUIPRESULT").Value & ""
                sResult2 = RsLocal.Fields("RESULT").Value & ""
                strJudge = RsLocal.Fields("REFJUDGE").Value & ""
                
                '-- 장비결과적용
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                  
                '-- 결과저장용 키 가져오기
                If strOrdCd = "" Then
                    strOrdCd = GetSampleSubITEM(strBarcode, strTestCd, strHospDate, strChartNo)
                End If
                
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" And strOrdCd <> "" Then
                    strDate = Format(Now, "yyyy-mm-dd")
                    strTime = Format(Now, "hh:mm:ss")
                
                    '-- 서버저장
                    SQL = ""
                    SQL = SQL & "UPDATE ME_LABDAT                           " & vbCrLf
                    SQL = SQL & "   Set LABRESULT = '" & sResult & "'       " & vbCrLf  '검사결과
                    SQL = SQL & "     , LABENDDEP = '2'                     " & vbCrLf  '처리상태       2:접수, 3:결과입력
                    SQL = SQL & "     , LABRSTDTE = '" & strDate & "'       " & vbCrLf  '결과입력일자   YYYY-MM-DD
                    SQL = SQL & "     , LABRSTTIM = '" & strTime & "'       " & vbCrLf  '결과입력일자   YYYY-MM-DD
                    SQL = SQL & "     , LABRSTUID = '" & gHOSP.USERID & "'  " & vbCrLf  '결과입력ID
                    SQL = SQL & "     , LABRSTCOM = '" & gHOSP.MACHNM & "'  " & vbCrLf  '결과입력컴퓨터명
                    SQL = SQL & " WHERE LABATTEND = '" & strPatID & "'      " & vbCrLf  '내원번호
                    SQL = SQL & "   And LABODRCOD = '" & strTestCd & "'     " & vbCrLf  '검사코드
                    SQL = SQL & "   And LABODRSTP = '" & strOrdCd & "'      " & vbCrLf  '검사일련번호
                    SQL = SQL & "   And LABODRDTE = '" & strHospDate & "'   " & vbCrLf
'                    SQL = SQL & "   And LABBARCOD = '" & strBarcode & "'    " & vbCrLf  '바코드
                    
                    Call SetSQLData("결과저장", SQL, "A")
                    AdoCn.Execute SQL
                                        
                    '-- 상태변경
                    SQL = ""
                    SQL = SQL & "UPDATE ME_DAT                              " & vbCrLf
                    SQL = SQL & "   Set DATENDDEP   = '2'                   " & vbCrLf  '처리상태       2:접수, 3:결과입력
                    SQL = SQL & "     , DATRSTDTE = '" & strDate & "'       " & vbCrLf  '결과입력일자   YYYY-MM-DD
                    SQL = SQL & "     , DATRSTTIM = '" & strTime & "'       " & vbCrLf  '결과입력시간   hh:mm:ss
                    SQL = SQL & "     , DATRSTUID = '" & gHOSP.USERID & "'  " & vbCrLf  '결과입력ID
                    SQL = SQL & "     , DATRSTCOM = '" & gHOSP.MACHNM & "'  " & vbCrLf  '결과입력컴퓨터명
                    SQL = SQL & " WHERE DATATTEND = '" & strPatID & "'      " & vbCrLf  '내원번호
                    SQL = SQL & "   And DATODRCOD = '" & strTestCd & "'     " & vbCrLf  '검사코드
                    SQL = SQL & "   And DATODRSTP = '" & strOrdCd & "'      " & vbCrLf  '검사일련번호
                    SQL = SQL & "   And DATODRDTE = '" & strHospDate & "'"
                    'SQL = SQL & "   And DATBARCOD = '" & strBarcode & "'    " & vbCrLf  '바코드
                    
                    Call SetSQLData("상태변경", SQL, "A")
                    AdoCn.Execute SQL
                    
                    blnSave = True
                            
                End If
                RsLocal.MoveNext
            Loop
        End If
        
        RsLocal.Close
        
'        If blnSave = True Then
'            '-- 상태변경
'            SQL = ""
'            SQL = SQL & "UPDATE ME_DAT Set " & vbCrLf
'            SQL = SQL & "   Set DATENDDEP   = '2' " & vbCrLf         '처리상태       2:접수, 3:결과입력
'            SQL = SQL & "     , DATRSTDTE = '" & strDate & "', " & vbCrLf  '결과입력일자   YYYY-MM-DD
'            SQL = SQL & "     , DATRSTTIM = '" & strTime & "', " & vbCrLf  '결과입력시간   hh:mm:ss
'            SQL = SQL & "     , DATRSTUID = '" & gHOSP.USERID & "', " & vbCrLf  '결과입력ID
'            SQL = SQL & "     , DATRSTCOM = '" & gHOSP.MACHNM & "' " & vbCrLf    '결과입력컴퓨터명
'            SQL = SQL & " WHERE DATATTEND = '" & strPatID & "'" & vbCrLf '내원번호
'            SQL = SQL & "   And DATODRCOD = " & gAllOrdCd & vbCrLf     '처방코드
'            SQL = SQL & "   And DATODRSTP = '" & strOrdCd & "'"       '검사일련번호
'            SQL = SQL & "   And DATODRDTE = '" & strHospDate & "'"
'            SQL = SQL & "   And DATBARCOD = '" & strBarcode & "'" & vbCr  '바코드
'
'            Call SetSQLData("상태변경", "최종상태변경 : " & SQL)
'
'            AdoCn.Execute SQL
'
'            SaveTransData_BIT70 = 1
'
'        End If
        SaveTransData_BIT70 = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_BIT70 = -1
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SaveTransData_BIT70" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

Function SaveTransData_BIT(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    Dim RsLocal         As ADODB.Recordset
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strPatNm        As String
    
    Dim strEqpcd        As String
    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strTestCdSub    As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strJudge        As String
    
    Dim strDate         As String
    Dim strTime         As String
    Dim blnSave         As Boolean
    
On Error GoTo ErrHandle
    
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""
    blnSave = False
    
    With frmInterface
        SaveTransData_BIT = -1
        
        strSaveSeq = Trim(GetText(SPD, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(SPD, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(SPD, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(SPD, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(SPD, argSpcRow, colPID))
        strPatNm = Trim(GetText(SPD, argSpcRow, colPNAME))
        strChartNo = Trim(GetText(SPD, argSpcRow, colCHARTNO))
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strChartNo) = "" Then
            Exit Function
        End If
        
        If Trim(strPatNm) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMCODESUB,EQUIPRESULT,RESULT,REFJUDGE    " & vbCrLf
        SQL = SQL & "  FROM PATRESULT                                                               " & vbCrLf
        SQL = SQL & " WHERE EXAMDATE    = '" & strExamDate & "'                                     " & vbCrLf
        SQL = SQL & "   AND SAVESEQ     = " & strSaveSeq & vbCrLf
        SQL = SQL & "   AND BARCODE     = '" & strBarcode & "'                                      " & vbCrLf
        SQL = SQL & "   AND EXAMCODE    <> ''                                                       " & vbCrLf
        
        Set RsLocal = New ADODB.Recordset
        Set RsLocal = AdoCn_Local.Execute(SQL, , 1)
        If Not RsLocal.EOF = True And Not RsLocal.BOF = True Then
            Do Until RsLocal.EOF
                strEqpcd = RsLocal.Fields("EQUIPCODE").Value & ""
                strOrdCd = RsLocal.Fields("ORDERCODE").Value & ""
                strTestCd = RsLocal.Fields("EXAMCODE").Value & ""
                strTestCdSub = RsLocal.Fields("EXAMCODESUB").Value & ""
                sResult1 = RsLocal.Fields("EQUIPRESULT").Value & ""
                sResult2 = RsLocal.Fields("RESULT").Value & ""
                strJudge = RsLocal.Fields("REFJUDGE").Value & ""
                
                '-- 장비결과적용
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                  
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                    strDate = Format(Now, "yyyy-mm-dd")
                    strTime = Format(Now, "hh:mm:ss")
                
                    '-- 서버저장
                    SQL = ""
                    SQL = SQL & "UPDATE RESINF                                                      " & vbCrLf
                    SQL = SQL & "   SET RESRLTVAL = '" & sResult & "'                               " & vbCrLf
                    SQL = SQL & "     , RESUPDDTM = '" & Format(Now, "yyyymmddhhmm") & "'           " & vbCrLf
                    SQL = SQL & "     , RESREPTYP = 'M'                                             " & vbCrLf       'M : 보고대기, N : 미결과, F : 보고
                    SQL = SQL & " WHERE LTRIM(RTRIM(ResOcmNum)) = '" & strBarcode & "'              " & vbCrLf
                    SQL = SQL & "   AND Upper(LTRIM(RTRIM(RESLABCOD))) = '" & UCase(strTestCd) & "' " & vbCrLf
                    SQL = SQL & "   AND (RESREPTYP = 'N' OR RESREPTYP IS NULL)                      " 'N : 미결과
                    
                    Call SetSQLData("결과저장", SQL, "A")
                    AdoCn.Execute SQL
                                        
                    blnSave = True
                            
                End If
                RsLocal.MoveNext
            Loop
        End If
        
        RsLocal.Close
        
        SaveTransData_BIT = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_BIT = -1
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SaveTransData_BIT" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

Function SaveTransData_MCC(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    Dim RsLocal         As ADODB.Recordset
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strPatNm        As String
    
    Dim strEqpcd        As String
    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strTestCdSub    As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strJudge        As String
    
    Dim dblBarno        As Double
    Dim Prm1            As New ADODB.Parameter
    Dim Prm2            As New ADODB.Parameter
    Dim Prm3            As New ADODB.Parameter
    Dim prm4            As New ADODB.Parameter
    Dim prm5            As New ADODB.Parameter
    
    Dim blnSave         As Boolean
    
    blnSave = False
    
On Error GoTo ErrHandle
    
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""

    With frmInterface
        SaveTransData_MCC = -1
        
        strSaveSeq = Trim(GetText(SPD, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(SPD, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(SPD, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(SPD, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(SPD, argSpcRow, colPID))
        strPatNm = Trim(GetText(SPD, argSpcRow, colPNAME))
        strChartNo = Trim(GetText(SPD, argSpcRow, colCHARTNO))
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatNm) = "" Then
            Exit Function
        End If
        
        If IsNumeric(strBarcode) Then
            dblBarno = CDbl(strBarcode)
        Else
            Exit Function
        End If
        

        '-- Local에서 환자별로 결과값 가져오기
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMCODESUB,EQUIPRESULT,RESULT,REFJUDGE    " & vbCrLf
        SQL = SQL & "  FROM PATRESULT                                                               " & vbCrLf
        SQL = SQL & " WHERE EXAMDATE    = '" & strExamDate & "'                                     " & vbCrLf
        SQL = SQL & "   AND SAVESEQ     = " & strSaveSeq & vbCrLf
        SQL = SQL & "   AND BARCODE     = '" & strBarcode & "'                                      " & vbCrLf
        SQL = SQL & "   AND EXAMCODE    <> ''                                                       " & vbCrLf
        
        Set RsLocal = New ADODB.Recordset
        Set RsLocal = AdoCn_Local.Execute(SQL, , 1)
        If Not RsLocal.EOF = True And Not RsLocal.BOF = True Then
            Do Until RsLocal.EOF
                strEqpcd = RsLocal.Fields("EQUIPCODE").Value & ""
                strOrdCd = RsLocal.Fields("ORDERCODE").Value & ""
                strTestCd = RsLocal.Fields("EXAMCODE").Value & ""
                strTestCdSub = RsLocal.Fields("EXAMCODESUB").Value & ""
                sResult1 = RsLocal.Fields("EQUIPRESULT").Value & ""
                sResult2 = RsLocal.Fields("RESULT").Value & ""
                strJudge = RsLocal.Fields("REFJUDGE").Value & ""
                
                '-- 장비결과적용
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                    '-- 서버저장
                    '-- 2020.01.16 : 중환자실용 SP 변경 (==> 자동확정 로직추가)

                    SQL = ""
                    SQL = SQL & "Exec UP_LIS_INTERFACE_U$014 " & dblBarno & "," & strTestCd & "," & sResult & "," & gHOSP.MACHCD & "," & gHOSP.USERID
            
                    Set AdoCmd = New ADODB.Command
                    Set AdoCmd.ActiveConnection = AdoCn
                    With AdoCmd
                        .CommandTimeout = 15
                        .CommandText = "UP_LIS_INTERFACE_U$014"
                        .CommandType = adCmdStoredProc
                        
                        Set Prm1 = .CreateParameter("BCODE_NO", adInteger, adParamInput, 30, dblBarno)      '바코드번호
                        .Parameters.Append Prm1
    
                        Set Prm2 = .CreateParameter("ORD_CD", adVarChar, adParamInput, 10, strTestCd)       '처방코드
                        .Parameters.Append Prm2
    
                        Set Prm3 = .CreateParameter("RESULT_NM", adVarChar, adParamInput, 4000, sResult)    '결과값
                        .Parameters.Append Prm3
    
                        Set prm4 = .CreateParameter("ENT_EMPL_NO", adVarChar, adParamInput, 15, gHOSP.USERID)     '장비코드 'B09' 또는 'B10'
                        .Parameters.Append prm4
                        
                        Set prm5 = .CreateParameter("EQP_CD", adVarChar, adParamInput, 15, gHOSP.MACHCD)    '장비코드 'B09' 또는 'B10'
                        .Parameters.Append prm5
    
                        .Execute
                        
                        blnSave = True

'                UP_LIS_INTERFACE_U$014 (@BCODE_NO      INT               --바코드번호
'               ,   @ORD_CD         VARCHAR(10)         --처방코드
'               ,   @RESULT_NM      VARCHAR(4000)      --결과값
'               ,   @ENT_EMPL_NO   VARCHAR(15)         --입력자 사번
'               ,   @EQP_CD         VARCHAR(15)   = NULL )--장비코드 기존과 동일
                        
                    End With
                    
                    Call SetSQLData("결과저장", SQL, "A")
                    
                End If
                RsLocal.MoveNext
            Loop
        End If
        
        RsLocal.Close
        
        If blnSave = True Then
            SaveTransData_MCC = 1
        End If
        
    End With

    
    
Exit Function

ErrHandle:
    SaveTransData_MCC = -1
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SaveTransData_MCC" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function


Function SaveTransData_TWIN(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    Dim RsLocal         As ADODB.Recordset
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strPatNm        As String
    
    Dim strEqpcd        As String
    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strTestCdSub    As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strJudge        As String
    
    Dim dblBarno        As Double
    Dim Prm1            As New ADODB.Parameter
    Dim Prm2            As New ADODB.Parameter
    Dim Prm3            As New ADODB.Parameter
    Dim prm4            As New ADODB.Parameter
    
On Error GoTo ErrHandle
    
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""

    With frmInterface
        SaveTransData_TWIN = -1
        
        strSaveSeq = Trim(GetText(SPD, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(SPD, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(SPD, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(SPD, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(SPD, argSpcRow, colPID))
        strPatNm = Trim(GetText(SPD, argSpcRow, colPNAME))
        strChartNo = Trim(GetText(SPD, argSpcRow, colCHARTNO))
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatNm) = "" Then
            Exit Function
        End If
        
        
        '-- Local에서 환자별로 결과값 가져오기
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMCODESUB,EQUIPRESULT,RESULT,REFJUDGE    " & vbCrLf
        SQL = SQL & "  FROM PATRESULT                                                               " & vbCrLf
        SQL = SQL & " WHERE EXAMDATE    = '" & strExamDate & "'                                     " & vbCrLf
        SQL = SQL & "   AND SAVESEQ     = " & strSaveSeq & vbCrLf
        SQL = SQL & "   AND BARCODE     = '" & strBarcode & "'                                      " & vbCrLf
        SQL = SQL & "   AND EXAMCODE    <> ''                                                       " & vbCrLf
        
        Set RsLocal = New ADODB.Recordset
        Set RsLocal = AdoCn_Local.Execute(SQL, , 1)
        If Not RsLocal.EOF = True And Not RsLocal.BOF = True Then
            Do Until RsLocal.EOF
                strEqpcd = RsLocal.Fields("EQUIPCODE").Value & ""
                strOrdCd = RsLocal.Fields("ORDERCODE").Value & ""
                strTestCd = RsLocal.Fields("EXAMCODE").Value & ""
                strTestCdSub = RsLocal.Fields("EXAMCODESUB").Value & ""
                sResult1 = RsLocal.Fields("EQUIPRESULT").Value & ""
                sResult2 = RsLocal.Fields("RESULT").Value & ""
                strJudge = RsLocal.Fields("REFJUDGE").Value & ""
                
                '-- 장비결과적용
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                    '-- 서버저장
                    SQL = ""
                    SQL = SQL & "Update TW_HSP_OCS.TWEXAM_RESULTC           " & vbCrLf
                    SQL = SQL & "   Set STATUS      = '4'                   " & vbCrLf  '검사상태
                    SQL = SQL & "     , RESULT      = '" & sResult & "'     " & vbCrLf  '검사결과
                    SQL = SQL & "     , RESULTDATE  = TRUNC(SYSDATE)        " & vbCrLf  '검사전송시간
                    SQL = SQL & " Where SPECNO      = '" & strBarcode & "'  " & vbCrLf  '검체번호
'                    SQL = SQL & "   And MASTERCODE  = 'LH1P01'    " & vbCrLf  '마스터코드 LH1P01
                    SQL = SQL & "   And SUBCODE     = '" & strTestCd & "'   " & vbCrLf  '검사코드
                    SQL = SQL & "   And STATUS      <= '3'                  " & vbCrLf  '검사상태(=검체접수)
                    
                    Call SetSQLData("결과저장", SQL, "A")
                    AdoCn.Execute SQL
                
                    '-- 상태업데이트
                    SQL = ""
                    SQL = SQL & "Update TW_HSP_OCS.TWEXAM_SPECMST           " & vbCrLf
                    SQL = SQL & "   Set STATUS     = '3'                    " & vbCrLf '검사상태 [결과등록(3:결과미확인, 4:부분전송)]
                    SQL = SQL & "     , RESULTDATE = TRUNC(SYSDATE)         " & vbCrLf
                    SQL = SQL & " Where SPECNO     = '" & strBarcode & "'   " & vbCrLf '검체번호
                    SQL = SQL & "   And STATUS     <= '3'                   " & vbCrLf '검사상태 [3:검체접수]
                    
                    Call SetSQLData("상태저장", SQL, "A")
                    AdoCn.Execute SQL
                    
                    
                End If
                RsLocal.MoveNext
            Loop
        End If
        
        RsLocal.Close
        
        SaveTransData_TWIN = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_TWIN = -1
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "SaveTransData_TWIN" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function


'-- 검사마스터 조회
Public Sub GetTestList()
    Dim intRow          As Long

    intRow = 0
    gAllTestCd = ""
    Erase gArrEQP

    SQL = ""
    SQL = SQL & "SELECT "
    SQL = SQL & "  a.SEQNO          " & vbCrLf
    SQL = SQL & ", a.SENDCHANNEL    " & vbCrLf
    SQL = SQL & ", a.RSLTCHANNEL    " & vbCrLf
    SQL = SQL & ", a.TESTNAME       " & vbCrLf
    SQL = SQL & ", a.ABBRNAME       " & vbCrLf
    SQL = SQL & ", a.RESPRECUSE     " & vbCrLf
    SQL = SQL & ", a.RESPREC        " & vbCrLf
    SQL = SQL & ", a.REFMLOW        " & vbCrLf
    SQL = SQL & ", a.REFMHIGH       " & vbCrLf
    SQL = SQL & ", a.REFFLOW        " & vbCrLf
    SQL = SQL & ", a.REFFHIGH       " & vbCrLf
    SQL = SQL & ", a.RESTYPE        " & vbCrLf
    SQL = SQL & ", a.CALYN          " & vbCrLf
    
'    SQL = SQL & " ,AMRLimit1,AMRLimit2,AMRLimit3,AMRLimit4,AMRLimit5        " & vbCrLf
'    SQL = SQL & " ,AMRResult1,AMRResult2,AMRResult3,AMRResult4,AMRResult5   " & vbCrLf
'    SQL = SQL & " ,AMRINResult                                              " & vbCrLf
    SQL = SQL & "  FROM EQPMASTER a           " & vbCrLf
    SQL = SQL & " WHERE a.EQUIPCD    = '" & gHOSP.MACHCD & "'" & vbCrLf
    SQL = SQL & " ORDER BY a.SEQNO ASC, a.TESTNAME "
    '-- Record Count 가져옴
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        '처방코드, SUB코드용 추가 16,17
        ReDim Preserve gArrEQP(AdoRs_Local.RecordCount, 17)

        Do Until AdoRs_Local.EOF
            intRow = intRow + 1
            gArrEQP(intRow, 1) = AdoRs_Local.Fields("SEQNO").Value & ""
            gArrEQP(intRow, 2) = "" 'AdoRs_Local.Fields("TESTCODE").Value & ""
            gArrEQP(intRow, 3) = AdoRs_Local.Fields("SENDCHANNEL").Value & ""
            gArrEQP(intRow, 4) = AdoRs_Local.Fields("RSLTCHANNEL").Value & ""
            gArrEQP(intRow, 5) = AdoRs_Local.Fields("TESTNAME").Value & ""
            gArrEQP(intRow, 6) = AdoRs_Local.Fields("ABBRNAME").Value & ""
            gArrEQP(intRow, 7) = AdoRs_Local.Fields("RESPRECUSE").Value & ""
            gArrEQP(intRow, 8) = AdoRs_Local.Fields("RESPREC").Value & ""
            gArrEQP(intRow, 9) = AdoRs_Local.Fields("REFMLOW").Value & ""
            gArrEQP(intRow, 10) = AdoRs_Local.Fields("REFMHIGH").Value & ""
            gArrEQP(intRow, 11) = AdoRs_Local.Fields("REFFLOW").Value & ""
            gArrEQP(intRow, 12) = AdoRs_Local.Fields("REFFHIGH").Value & ""
            gArrEQP(intRow, 13) = AdoRs_Local.Fields("RESTYPE").Value & ""
            '-- 계산항목 여부
            gArrEQP(intRow, 14) = AdoRs_Local.Fields("CALYN").Value & ""
            gArrEQP(intRow, 16) = ""    '처방코드로 사용
            gArrEQP(intRow, 17) = ""    'SUB코드로 사용

'            If Trim(AdoRs_Local.Fields("TESTCODE").Value) <> "" Then
'                If intRow = 1 Or gAllTestCd = "" Then
'                    gAllTestCd = "'" & AdoRs_Local.Fields("TESTCODE").Value & "'"
'                Else
'                    gAllTestCd = gAllTestCd & ",'" & AdoRs_Local.Fields("TESTCODE").Value & "'"
'                End If
'            End If

            AdoRs_Local.MoveNext
        Loop
    End If

End Sub

'-- 검사마스터 코드조회
Public Sub GetTestCodeList()
    Dim intRow          As Long

    intRow = 0
    gAllTestCd = ""

    SQL = ""
    SQL = SQL & "SELECT "
    SQL = SQL & "  b.TESTCODE     " & vbCrLf
    SQL = SQL & "  FROM TESTMASTER b            " & vbCrLf
    SQL = SQL & " ORDER BY b.TESTCODE "
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            intRow = intRow + 1
            If Trim(AdoRs_Local.Fields("TESTCODE").Value) <> "" Then
                If intRow = 1 Or gAllTestCd = "" Then
                    gAllTestCd = "'" & AdoRs_Local.Fields("TESTCODE").Value & "'"
                Else
                    gAllTestCd = gAllTestCd & ",'" & AdoRs_Local.Fields("TESTCODE").Value & "'"
                End If
            End If
            AdoRs_Local.MoveNext
        Loop
    End If

End Sub


'-- 검사마스터명 조회
Public Sub GetTestListName()
    Dim intRow          As Long

    intRow = 0
    Erase gArrEQPNm

    SQL = ""
    SQL = SQL & "SELECT DISTINCT SEQNO,SENDCHANNEL,RSLTCHANNEL,ABBRNAME " & vbCrLf
'    SQL = SQL & " ,RESPRECUSE,RESPREC,REFMLOW,REFMHIGH,REFFLOW,REFFHIGH " & vbCrLf
    SQL = SQL & "  FROM EQPMASTER " & vbCr
    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "'" & vbCr
    SQL = SQL & " ORDER BY SEQNO "

    '-- Record Count 가져옴
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        
        ReDim Preserve gArrEQPNm(AdoRs_Local.RecordCount, 12)

        Do Until AdoRs_Local.EOF
            intRow = intRow + 1
            gArrEQPNm(intRow, 1) = AdoRs_Local.Fields("SEQNO").Value & ""
            gArrEQPNm(intRow, 2) = ""
            gArrEQPNm(intRow, 3) = AdoRs_Local.Fields("SENDCHANNEL").Value & ""
            gArrEQPNm(intRow, 4) = AdoRs_Local.Fields("RSLTCHANNEL").Value & ""
            gArrEQPNm(intRow, 5) = ""
            gArrEQPNm(intRow, 6) = AdoRs_Local.Fields("ABBRNAME").Value & ""
'            gArrEQPNm(intRow, 7) = AdoRs_Local.Fields("RESPRECUSE").Value & ""
'            gArrEQPNm(intRow, 8) = AdoRs_Local.Fields("RESPREC").Value & ""
'            gArrEQPNm(intRow, 9) = AdoRs_Local.Fields("REFMLOW").Value & ""
'            gArrEQPNm(intRow, 10) = AdoRs_Local.Fields("REFMHIGH").Value & ""
'            gArrEQPNm(intRow, 11) = AdoRs_Local.Fields("REFFLOW").Value & ""
'            gArrEQPNm(intRow, 12) = AdoRs_Local.Fields("REFFHIGH").Value & ""

            AdoRs_Local.MoveNext
        Loop
    End If

End Sub


'-- 검사마스터 조회
Public Sub GetTestMaster(ByVal SPD As Object)
    Dim intRow          As Long
    Dim i               As Integer
    
    SPD.Font = "맑은 고딕"
    SPD.FontSize = 9
    SPD.FontBold = False
    
    SPD.MaxRows = 0
    intRow = 0
    
    For i = 0 To 3
        frmTestSet.cboCalTest(i).Clear
    Next
                
    SQL = ""
    SQL = SQL & "SELECT DISTINCT    "
    SQL = SQL & "  a.EQUIPCD        AS EQUIPCD      " & vbCrLf
    SQL = SQL & ", a.SEQNO          AS SEQNO        " & vbCrLf
    SQL = SQL & ", a.SENDCHANNEL    " & vbCrLf
    SQL = SQL & ", a.RSLTCHANNEL    AS RSLTCHANNEL  " & vbCrLf
    SQL = SQL & ", (SELECT TOP 1 b.TESTCODE FROM TESTMASTER b WHERE b.RSLTCHANNEL = a.RSLTCHANNEL) AS TESTCODE   " & vbCrLf
    SQL = SQL & ", a.TESTNAME       " & vbCrLf
    SQL = SQL & ", a.ABBRNAME       " & vbCrLf
    SQL = SQL & ", a.RESPRECUSE     " & vbCrLf
    SQL = SQL & ", a.RESPREC        " & vbCrLf
    SQL = SQL & ", a.REFMLOW        " & vbCrLf
    SQL = SQL & ", a.REFMHIGH       " & vbCrLf
    SQL = SQL & ", a.REFFLOW        " & vbCrLf
    SQL = SQL & ", a.REFFHIGH       " & vbCrLf
    SQL = SQL & ", a.RESTYPE        " & vbCrLf
    SQL = SQL & ", a.CALYN          " & vbCrLf
    SQL = SQL & ", b.AMRINResult    " & vbCrLf
    SQL = SQL & ", b.AMRLimit1,     b.AMRLimit2,    b.AMRLimit3,    b.AMRLimit4,    b.AMRLimit5     " & vbCrLf
    SQL = SQL & ", b.AMRLimit6,     b.AMRLimit7,    b.AMRLimit8,    b.AMRLimit9,    b.AMRLimit10    " & vbCrLf
    SQL = SQL & ", b.AMRResult1,    b.AMRResult2,   b.AMRResult3,   b.AMRResult4,   b.AMRResult5    " & vbCrLf
    SQL = SQL & ", b.AMRResult6,    b.AMRResult7,   b.AMRResult8,   b.AMRResult9,   b.AMRResult10   " & vbCrLf
    
    SQL = SQL & "  FROM (EQPMASTER a LEFT OUTER JOIN AMRMASTER b   " & vbCrLf
    SQL = SQL & "    ON a.RSLTCHANNEL   = b.RSLTCHANNEL)            " & vbCrLf
'    SQL = SQL & "  LEFT OUTER JOIN TESTMASTER c                      " & vbCrLf
'    SQL = SQL & "    ON a.RSLTCHANNEL   = c.RSLTCHANNEL             " & vbCrLf
    SQL = SQL & " WHERE a.EQUIPCD       = '" & gHOSP.MACHCD & "'    " & vbCrLf
    SQL = SQL & " ORDER BY a.SEQNO "

    '-- Record Count 가져옴
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        With SPD
            .MaxRows = AdoRs_Local.RecordCount '

            Do Until AdoRs_Local.EOF
                intRow = intRow + 1
                Call SetText(SPD, AdoRs_Local.Fields("EQUIPCD").Value & "", intRow, colLMACHCODE)
                Call SetText(SPD, AdoRs_Local.Fields("SEQNO").Value & "", intRow, colLSEQNO)
                Call SetText(SPD, AdoRs_Local.Fields("TESTCODE").Value & "", intRow, colLTESTCD)
                Call SetText(SPD, AdoRs_Local.Fields("SENDCHANNEL").Value & "", intRow, colLOCHANNEL)
                Call SetText(SPD, AdoRs_Local.Fields("RSLTCHANNEL").Value & "", intRow, colLRCHANNEL)
                Call SetText(SPD, AdoRs_Local.Fields("TESTNAME").Value & "", intRow, colLTESTNM)
                Call SetText(SPD, AdoRs_Local.Fields("ABBRNAME").Value & "", intRow, colLABBRNM)
                Call SetText(SPD, IIf(AdoRs_Local.Fields("RESPRECUSE").Value & "" = "1", "1", "0"), intRow, colLRESSPECUSE)
                Call SetText(SPD, AdoRs_Local.Fields("RESPREC").Value & "", intRow, colLRESSPEC)
                Call SetText(SPD, AdoRs_Local.Fields("REFMLOW").Value & "", intRow, colLMLOW)
                Call SetText(SPD, AdoRs_Local.Fields("REFMHIGH").Value & "", intRow, colLMHIGH)
                Call SetText(SPD, AdoRs_Local.Fields("REFFLOW").Value & "", intRow, colLFLOW)
                Call SetText(SPD, AdoRs_Local.Fields("REFFHIGH").Value & "", intRow, colLFHIGH)
                If AdoRs_Local.Fields("RESTYPE").Value & "" = "0" Then
                    Call SetText(SPD, "수치형", intRow, colRESTYPE)
                ElseIf AdoRs_Local.Fields("RESTYPE").Value & "" = "1" Then
                    Call SetText(SPD, "문자형", intRow, colRESTYPE)
                ElseIf AdoRs_Local.Fields("RESTYPE").Value & "" = "2" Then
                    Call SetText(SPD, "수치/문자형", intRow, colRESTYPE)
                Else
                    Call SetText(SPD, AdoRs_Local.Fields("RESTYPE").Value & "", intRow, colRESTYPE)
                End If
                
                Call SetText(SPD, AdoRs_Local.Fields("AMRLimit1").Value & "|" & AdoRs_Local.Fields("AMRResult1").Value & "", intRow, colRESTYPE + 1)
                Call SetText(SPD, AdoRs_Local.Fields("AMRLimit2").Value & "|" & AdoRs_Local.Fields("AMRResult2").Value & "", intRow, colRESTYPE + 2)
                Call SetText(SPD, AdoRs_Local.Fields("AMRLimit3").Value & "|" & AdoRs_Local.Fields("AMRResult3").Value & "", intRow, colRESTYPE + 3)
                Call SetText(SPD, AdoRs_Local.Fields("AMRLimit4").Value & "|" & AdoRs_Local.Fields("AMRResult4").Value & "", intRow, colRESTYPE + 4)
                Call SetText(SPD, AdoRs_Local.Fields("AMRLimit5").Value & "|" & AdoRs_Local.Fields("AMRResult5").Value & "", intRow, colRESTYPE + 5)
                Call SetText(SPD, AdoRs_Local.Fields("AMRLimit6").Value & "|" & AdoRs_Local.Fields("AMRResult6").Value & "", intRow, colRESTYPE + 6)
                Call SetText(SPD, AdoRs_Local.Fields("AMRLimit7").Value & "|" & AdoRs_Local.Fields("AMRResult7").Value & "", intRow, colRESTYPE + 7)
                If AdoRs_Local.Fields("AMRINResult").Value & "" = "0" Then
                    Call SetText(SPD, "변환없음", intRow, colRESTYPE + 8)
                ElseIf AdoRs_Local.Fields("AMRINResult").Value & "" = "1" Then
                    Call SetText(SPD, "판정(수치)", intRow, colRESTYPE + 8)
                ElseIf AdoRs_Local.Fields("AMRINResult").Value & "" = "2" Then
                    Call SetText(SPD, "수치(판정)", intRow, colRESTYPE + 8)
                ElseIf AdoRs_Local.Fields("AMRINResult").Value & "" = "3" Then
                    Call SetText(SPD, "판정 수치", intRow, colRESTYPE + 8)
                ElseIf AdoRs_Local.Fields("AMRINResult").Value & "" = "4" Then
                    Call SetText(SPD, "수치 판정", intRow, colRESTYPE + 8)
                End If
                
                Call SetText(SPD, AdoRs_Local.Fields("CALYN").Value & "", intRow, colRESTYPE + 9)
                
                
                For i = 0 To 3
                    frmTestSet.cboCalTest(i).AddItem AdoRs_Local.Fields("TESTNAME").Value & ""
                Next
                
                AdoRs_Local.MoveNext
            Loop
            .RowHeight(-1) = 15
        End With
    End If



End Sub


'-- AMR마스터 조회
Public Sub GetAMRMaster(ByVal pSeqNo As Integer, ByVal pRCd As String, ByVal pTestCd As String)

    SQL = ""
    SQL = SQL & "SELECT * " & vbCrLf
    SQL = SQL & "  FROM AMRMASTER " & vbCr
    SQL = SQL & " WHERE EQUIPCD   = '" & gHOSP.MACHCD & "'" & vbCrLf
    SQL = SQL & "   AND RSLTCHANNEL  = '" & pRCd & "'" & vbCrLf

    '-- Record Count 가져옴
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            frmTestSet.txtAMRLimit(0).Text = AdoRs_Local.Fields("AMRLIMIT1").Value & ""
            frmTestSet.txtAMRLimit(1).Text = AdoRs_Local.Fields("AMRLIMIT2").Value & ""
            frmTestSet.txtAMRLimit(2).Text = AdoRs_Local.Fields("AMRLIMIT3").Value & ""
            frmTestSet.txtAMRLimit(3).Text = AdoRs_Local.Fields("AMRLIMIT4").Value & ""
            frmTestSet.txtAMRLimit(4).Text = AdoRs_Local.Fields("AMRLIMIT5").Value & ""
            frmTestSet.txtAMRLimit(5).Text = AdoRs_Local.Fields("AMRLIMIT6").Value & ""
            frmTestSet.txtAMRLimit(6).Text = AdoRs_Local.Fields("AMRLIMIT7").Value & ""
            '문자형
            frmTestSet.txtAMRLimit(7).Text = AdoRs_Local.Fields("AMRLIMIT8").Value & ""
            frmTestSet.txtAMRLimit(8).Text = AdoRs_Local.Fields("AMRLIMIT9").Value & ""
            frmTestSet.txtAMRLimit(9).Text = AdoRs_Local.Fields("AMRLIMIT10").Value & ""
            frmTestSet.txtAMRLimit(10).Text = AdoRs_Local.Fields("AMRLIMIT11").Value & ""
            frmTestSet.txtAMRLimit(11).Text = AdoRs_Local.Fields("AMRLIMIT12").Value & ""
            frmTestSet.txtAMRLimit(12).Text = AdoRs_Local.Fields("AMRLIMIT13").Value & ""
            frmTestSet.txtAMRLimit(13).Text = AdoRs_Local.Fields("AMRLIMIT14").Value & ""
            '수치형
            frmTestSet.txtAMRLimit(14).Text = AdoRs_Local.Fields("AMRLIMIT15").Value & ""
            frmTestSet.txtAMRLimit(15).Text = AdoRs_Local.Fields("AMRLIMIT16").Value & ""
            frmTestSet.txtAMRLimit(16).Text = AdoRs_Local.Fields("AMRLIMIT17").Value & ""
            frmTestSet.txtAMRLimit(17).Text = AdoRs_Local.Fields("AMRLIMIT18").Value & ""
            frmTestSet.txtAMRLimit(18).Text = AdoRs_Local.Fields("AMRLIMIT19").Value & ""
            
            frmTestSet.txtAMRResult(0).Text = AdoRs_Local.Fields("AMRRESULT1").Value & ""
            frmTestSet.txtAMRResult(1).Text = AdoRs_Local.Fields("AMRRESULT2").Value & ""
            frmTestSet.txtAMRResult(2).Text = AdoRs_Local.Fields("AMRRESULT3").Value & ""
            frmTestSet.txtAMRResult(3).Text = AdoRs_Local.Fields("AMRRESULT4").Value & ""
            frmTestSet.txtAMRResult(4).Text = AdoRs_Local.Fields("AMRRESULT5").Value & ""
            frmTestSet.txtAMRResult(5).Text = AdoRs_Local.Fields("AMRRESULT6").Value & ""
            frmTestSet.txtAMRResult(6).Text = AdoRs_Local.Fields("AMRRESULT7").Value & ""
            frmTestSet.txtAMRResult(7).Text = AdoRs_Local.Fields("AMRRESULT8").Value & ""
            frmTestSet.txtAMRResult(8).Text = AdoRs_Local.Fields("AMRRESULT9").Value & ""
            frmTestSet.txtAMRResult(9).Text = AdoRs_Local.Fields("AMRRESULT10").Value & ""
            frmTestSet.txtAMRResult(10).Text = AdoRs_Local.Fields("AMRRESULT11").Value & ""
            frmTestSet.txtAMRResult(11).Text = AdoRs_Local.Fields("AMRRESULT12").Value & ""
            frmTestSet.txtAMRResult(12).Text = AdoRs_Local.Fields("AMRRESULT13").Value & ""
            frmTestSet.txtAMRResult(13).Text = AdoRs_Local.Fields("AMRRESULT14").Value & ""
            frmTestSet.txtAMRResult(14).Text = AdoRs_Local.Fields("AMRRESULT15").Value & ""
            frmTestSet.txtAMRResult(15).Text = AdoRs_Local.Fields("AMRRESULT16").Value & ""
            frmTestSet.txtAMRResult(16).Text = AdoRs_Local.Fields("AMRRESULT17").Value & ""
            frmTestSet.txtAMRResult(17).Text = AdoRs_Local.Fields("AMRRESULT18").Value & ""
            frmTestSet.txtAMRResult(18).Text = AdoRs_Local.Fields("AMRRESULT19").Value & ""
            
            AdoRs_Local.MoveNext
        Loop
    End If

End Sub

'-- 검사오더마스터 조회
Public Function GetTestCode(ByVal pChannel As String) As String
    Dim strTestCode As String
    
    GetTestCode = ""
    strTestCode = ""

    SQL = ""
    SQL = SQL & "SELECT DISTINCT TESTCODE "
    SQL = SQL & "  FROM TESTMASTER " & vbCrLf
    SQL = SQL & " WHERE RSLTCHANNEL = '" & pChannel & "'" & vbCrLf
    SQL = SQL & " ORDER BY TESTCODE "

    '-- Record Count 가져옴
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strTestCode = strTestCode & AdoRs_Local.Fields("TESTCODE").Value & "@"
            AdoRs_Local.MoveNext
        Loop
    End If
    
    GetTestCode = strTestCode

End Function

'-- 계산항목조회
Public Function GetCalContents(ByVal pChannel As String, ByVal pCode As String) As String
    Dim strCalContents As String
    
    GetCalContents = ""
    strCalContents = ""

    SQL = ""
    SQL = SQL & "SELECT DISTINCT CALCULATE "
    SQL = SQL & "  FROM TESTMASTER " & vbCrLf
    SQL = SQL & " WHERE RSLTCHANNEL = '" & pChannel & "'" & vbCrLf
    If pCode <> "" Then
        SQL = SQL & "   AND TESTCODE    = '" & pCode & "'" & vbCrLf
    End If
    SQL = SQL & "   AND CALCULATE <> '' "
    '-- Record Count 가져옴
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strCalContents = AdoRs_Local.Fields("CALCULATE").Value & ""
            AdoRs_Local.MoveNext
        Loop
    End If
    
    GetCalContents = strCalContents

End Function

'-- 검사명으로 채널찾기
Public Function GetChannel(ByVal pTestName As String) As String
    Dim strChannel As String
    
    GetChannel = ""
    strChannel = ""

    SQL = ""
    SQL = SQL & "SELECT DISTINCT RSLTCHANNEL "
    SQL = SQL & "  FROM EQPMASTER " & vbCrLf
    SQL = SQL & " WHERE TESTNAME = '" & pTestName & "'" & vbCrLf

    '-- Record Count 가져옴
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strChannel = AdoRs_Local.Fields("RSLTCHANNEL").Value & ""
            AdoRs_Local.MoveNext
        Loop
    End If
    
    GetChannel = strChannel

End Function


Public Function GetMaxSeqCode() As Integer
    Dim strTestCode As String
    
    GetMaxSeqCode = 0
    
    SQL = ""
    SQL = SQL & "SELECT MAX(SEQNO) AS MAXSEQ "
    SQL = SQL & "  FROM EQPMASTER " & vbCrLf
'    SQL = SQL & " WHERE RSLTCHANNEL = '" & pChannel & "'" & vbCrLf
'    SQL = SQL & " ORDER BY TESTCODE "

    '-- Record Count 가져옴
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            GetMaxSeqCode = AdoRs_Local.Fields("MAXSEQ").Value
            AdoRs_Local.MoveNext
        Loop
    End If
    
    If GetMaxSeqCode = 0 Then
        GetMaxSeqCode = 1
    Else
        GetMaxSeqCode = GetMaxSeqCode + 1
    End If
    
End Function

'-- 검사명으로 채널 조회
Public Function GetTestCh(ByVal pItem As String) As String
    Dim intRow          As Long

    GetTestCh = ""

    SQL = ""
    SQL = SQL & "SELECT RSLTCHANNEL                 " & vbCrLf
    SQL = SQL & "  FROM EQPMASTER                   " & vbCrLf
    SQL = SQL & " WHERE ABBRNAME = '" & pItem & "'  " & vbCrLf
        
    '-- Record Count 가져옴
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            GetTestCh = AdoRs_Local.Fields("RSLTCHANNEL").Value & ""
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close

End Function

'-- 검사코드로 검사명 조회
Public Function GetTestNm(ByVal pItem As String, Optional pFull As Boolean) As String
    Dim intRow          As Long

    GetTestNm = ""

    If pFull = True Then
        SQL = ""
        SQL = SQL & "SELECT a.TESTNAME AS ITEMNM            " & vbCrLf
        SQL = SQL & "  FROM EQPMASTER a, TESTMASTER b       " & vbCrLf
        SQL = SQL & " WHERE a.RSLTCHANNEL = b.RSLTCHANNEL   " & vbCrLf
        SQL = SQL & "   AND b.TESTCODE = '" & pItem & "'    " & vbCrLf
    Else
        SQL = ""
        SQL = SQL & "SELECT a.ABBRNAME AS ITEMNM            " & vbCrLf
        SQL = SQL & "  FROM EQPMASTER a, TESTMASTER b       " & vbCrLf
        SQL = SQL & " WHERE a.RSLTCHANNEL = b.RSLTCHANNEL   " & vbCrLf
        SQL = SQL & "   AND b.TESTCODE = '" & pItem & "'    " & vbCrLf
    End If

    '-- Record Count 가져옴
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            GetTestNm = AdoRs_Local.Fields("ITEMNM").Value & ""
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close

End Function

'-- 검사코드로 검사명 조회
Public Function GetTestNmS(ByVal pItem As String, Optional pFull As Boolean) As String
    Dim strTestNM   As String
    
    strTestNM = ""
    GetTestNmS = ""

    If pFull = True Then
        SQL = ""
        SQL = SQL & "SELECT TESTNAME AS ITEMNM FROM EQPMASTER " & vbCr
        SQL = SQL & " WHERE TESTCODE IN (" & STS(pItem) & ")"
    Else
        SQL = ""
        SQL = SQL & "SELECT ABBRNAME AS ITEMNM FROM EQPMASTER " & vbCr
        SQL = SQL & " WHERE TESTCODE IN (" & STS(pItem) & ")"
    End If

    '-- Record Count 가져옴
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strTestNM = strTestNM & AdoRs_Local.Fields("ITEMNM").Value & "/"
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close
    
    GetTestNmS = strTestNM

End Function

'
''-- 검사명으로 결과채널 조회
'Public Function GetRsltChannel(ByVal pItem As String) As String
'    Dim RS2             As ADODB.Recordset
'    Dim intRow          As Long
'
'    GetRsltChannel = ""
'
'    SQL = ""
'    SQL = SQL & "SELECT RSLTCHANNEL "
'    SQL = SQL & "  FROM EQPMASTER " & vbCr
'    SQL = SQL & " WHERE ABBRNAME = '" & pItem & "'"
'
'    Set RS2 = New ADODB.Recordset
'
'    '-- Record Count 가져옴
'    AdoCn_Local.CursorLocation = adUseClient
'    Set RS2 = AdoCn_Local.Execute(SQL, , 1)
'    If Not RS2.EOF = True And Not RS2.BOF = True Then
'        Do Until RS2.EOF
'            GetRsltChannel = RS2.Fields("RSLTCHANNEL").Value & ""
'            RS2.MoveNext
'        Loop
'    End If
'
'    RS2.Close
'
'End Function
'
''-- 검사항목 조회
'Public Function GetTest(ByVal pTestCd As String) As String
'
'On Error GoTo RST
'    GetTest = ""
'
'    SQL = ""
'    SQL = SQL & "Select ORD_NM "
'    SQL = SQL & "  From LIS_ORD_LIST_V" & vbCr
'    SQL = SQL & " Where ord_cd = '" & pTestCd & "'" & vbCr
'
'    '-- Record Count 가져옴
'    AdoCn.CursorLocation = adUseClient
'    Set RS = AdoCn.Execute(SQL, , 1)
'    If Not RS.EOF = True And Not RS.BOF = True Then
'        Do Until RS.EOF
'            GetTest = Trim(RS.Fields("ORD_NM")) & ""
'            RS.MoveNext
'        Loop
'    End If
'
'    RS.Close
'
'Exit Function
'
'RST:
'
'                strErrMsg = "위    치 : " & gHOSP.MACHNM & "GetTest" & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
'    frmErrMsg.txtErr = vbNewLine & strErrMsg
'    frmErrMsg.Show 'vbModal
'
'    Screen.MousePointer = 0
'
'End Function
'
''-- 워크리스트 조회
Public Sub GetWorkList(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As Object, Optional ByVal pFromNo As String = "", Optional ByVal pToNo As String = "", Optional ByVal pSave As Boolean = False)

    Select Case gEMR
        Case "UBCARE"                       '의사랑
                Call GetWorkList_UBCARE(pFrom, pTo, SPD)
        
        Case "BIT"                                '비트
                Call GetWorkList_BIT(pFrom, pTo, SPD)
        
        Case "EASYS"                        '이지스, MCC
                Call GetWorkList_EASYS(pFrom, pTo, SPD, pSave)
        
        Case "HEALTH"                       '보건소
'                Call GetWorkList_HEALTH(pFrom, pTo, SPD)
        
        Case "BITJSON"                                '비트 JSON
                Call GetWorkList_BITJSON(pFrom, pTo, SPD, pFromNo, pToNo)
    
        Case "NEOSENSE"                    '네오소프트(SENSE)
                Call GetWorkList_NEOSENSE(pFrom, pTo, SPD)
        
        Case "MCC"                          'MCC SP버전
                Call GetWorkList_MCC(pFrom, pTo, SPD)
        
        Case "JWINFO"
                Call GetWorkList_JWINFO(pFrom, pTo, SPD)
        
        Case "AMIS"                                 '아미스 테크놀러지
                Call GetWorkList_AMIS(pFrom, pTo, SPD)
        
        Case "EONM"                                 '이온엠
                Call GetWorkList_EONM(pFrom, pTo, SPD)

        Case "BIT70"                                '비트
                Call GetWorkList_BIT70(pFrom, pTo, SPD)

        Case "LABSPEAR"                             '이노베스트(필의료재단)
                Call GetWorkList_LABSPEAR(pFrom, pTo, SPD)

        Case "SANSOFT"                              '테스트서버
                Call GetWorkList_LABSPEAR(pFrom, pTo, SPD)

        Case "MEDICHART"                            '메디챠트
                Call GetWorkList_MEDICHART(pFrom, pTo, SPD)

        Case "KCHART"                               '다대소프트
                Call GetWorkList_KCHART(pFrom, pTo, SPD)

        Case "SY"
                Call GetWorkList_SY(pFrom, pTo, SPD)

'        Case "PHILL"
'                Call GetWorkList_PHILL(pFrom, pTo, SPD)
'
'        Case "MSINFOTEC"                    'MS인포텍
'                Call GetWorkList_MSINFOTEC(pFrom, pTo, SPD)
'
'        Case "HANARO"                       '하나로의료재단
'                Call GetWorkList_HANARO(pFrom, pTo, SPD)

'        Case "AMIS"                         '아미스
'                Call GetWorkList_AMIS(pFrom, pTo, SPD)
'
'        Case "BIGUBCARE"
'                Call GetWorkList_BIGUBCARE(pFrom, pTo, SPD)
'
'        Case "BIT"                          '비트
'                Call GetWorkList_BIT(pFrom, pTo, SPD)
'
'        Case "BIT70"                        '비트 HIB70
'                Call GetWorkList_BIT70(pFrom, pTo, SPD)
'
'        Case "EMEDI"                        '이메디
'                Call GetWorkList_AMIS(pFrom, pTo, SPD)
'
'
''        Case "EONM"                         '이온엠
''                Call GetWorkList_EONM(pFrom, pTo, SPD)
'
'        Case "GINUS"                         '지누스
''                Call GetWorkList_GINUS(pFrom, pTo, SPD)
'
'        Case "GSEN"                         '지센커뮤니케이션즈(이챠트)
'                Call GetWorkList_MSINFOTEC(pFrom, pTo, SPD)
'
'        Case "HWASAN"                       '화산
'                Call GetWorkList_HWASAN(pFrom, pTo, SPD)
'
'        Case "JAINCOM"                      '자인컴
'                Call GetWorkList_JAINCOM(pFrom, pTo, SPD)
'
'        Case "JWINFO"                       '중외정보
'                Call GetWorkList_JWINFO(pFrom, pTo, SPD)
'
'
'        Case "KOMAIN"                       '중외정보
'                Call GetWorkList_KOMAIN(pFrom, pTo, SPD)
'
'        Case "KYU"                          '건양대학교병원 - 워크리스트 기능없음
'                'Call GetWorkList_KYU(pFrom, pTo, SPD)
'

'        Case "MEDIIT"                       '메디IT(필의료재단)
'                Call GetWorkList_MEDIIT(pFrom, pTo, SPD)
'
'        Case "MEDITOLISS"                   '아름누리
'                Call GetWorkList_MEDITOLISS(pFrom, pTo, SPD)
'
'        Case "MCC"                          'MCC SP버전
'                Call GetWorkList_MCC(pFrom, pTo, SPD)
'
'        Case "MOD"                          'MOD 시스템
'                Call GetWorkList_MOD(pFrom, pTo, SPD)
'
'        Case "MSINFOTEC"                    'MS인포텍
'                Call GetWorkList_MSINFOTEC(pFrom, pTo, SPD)
'
'        Case "NEOSOFT"                      '네오소프트
'                Call GetWorkList_NEOSOFT(pFrom, pTo, SPD)
'
'        Case "ONITGUM"                      '온아티 검진
'                Call GetWorkList_ONITGUM(pFrom, pTo, SPD)
'
'        Case "ONITEMR"                      '온아티 EMR
'                Call GetWorkList_ONITEMR(pFrom, pTo, SPD)
'
'        Case "PLIS"                         '포미스 슈바이처
'                Call GetWorkList_PLIS(pFrom, pTo, SPD)
'
'        Case "SY"                           'SY
'                Call GetWorkList_SY(Format(pFrom, "yyyy-mm-dd"), Format(pTo, "yyyy-mm-dd"), SPD)
'
'        Case "TWIN"                         '투윈정보
'                Call GetWorkList_TWIN(pFrom, pTo, SPD)
'

'        Case "WELL"                         '웰커머스
'                Call GetWorkList_WELL(pFrom, pTo, SPD)

'        Case "ONIT"
'            Call GetWorkList_onit(pFrom, pTo, SPD)

'        Case "PLIS"
'            Call GetWorkList_PLIS(pFrom, pTo, SPD)
        Case Else


    End Select

End Sub

'필의료재단 OLD 버전
Public Sub GetWorkList_PHILL(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As Object)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean

    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strTestCds  As String
    
On Error GoTo ErrHandle

    Screen.MousePointer = 11
    blnSame = False
    strTestCds = ""

    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "   P.request_date  AS HOSPDATE " & vbCrLf
    SQL = SQL & " , P.exam_no       AS PID      " & vbCrLf
    SQL = SQL & " , P.company_code  AS DEPT     " & vbCrLf
    SQL = SQL & " , P.chart_no      AS CHARTNO  " & vbCrLf
    SQL = SQL & " , p.personal_id   AS BARCODE  " & vbCrLf
    SQL = SQL & " , p.person_name   AS PNAME    " & vbCrLf
    SQL = SQL & " , P.worker_code               " & vbCrLf
    SQL = SQL & " , P.patient_kind              " & vbCrLf
    SQL = SQL & " , P.person_sex    AS SEX      " & vbCrLf
    SQL = SQL & " , P.person_age    AS AGE      " & vbCrLf
    SQL = SQL & " , R.pro_code      AS ITEM     " & vbCrLf
    SQL = SQL & "  FROM trust P, trures R       " & vbCrLf
    SQL = SQL & " WHERE P.request_date BETWEEN '" & pFrom & "' AND '" & pTo & "'" & vbCrLf
    SQL = SQL & "   AND R.pro_code IN (" & gAllTestCd & ") " & vbCrLf
    SQL = SQL & "   AND R.exam_code <> 'X999' " & vbCrLf
    SQL = SQL & "   AND P.request_date = R.request_date " & vbCrLf
    SQL = SQL & "   AND P.exam_no = R.exam_no " & vbCrLf
    SQL = SQL & " ORDER BY P.request_date, P.exam_no " & vbCrLf

    Call SetSQLData("워크조회", SQL, "")

    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then

        SPD.MaxRows = 0

        Do Until RS.EOF
            With SPD
                strTestCds = strTestCds & "'" & Trim(RS.Fields("ITEM")) & "',"

                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strBarcode = GetText(SPD, i, colBARCODE)
                    If Trim(RS("HOSPDATE")) = strHospDate And Trim(RS("BARCODE")) = strBarcode Then
                        blnSame = True
                    End If
                Next

                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows

                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                    SetText SPD, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
                    SetText SPD, Trim(RS.Fields("DEPT")) & "", intRow, colDEPT
                    
                    SetText SPD, GetTestNmS(Mid(strTestCds, 1, Len(strTestCds) - 1)), intRow, colSTATE + 1
                
                End If
                
            End With

            blnSame = False

            DoEvents

            RS.MoveNext
        Loop
    Else
        frmInterface.lblComStatus.Caption = "워크리스트 조회 대상자가 없습니다."
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetWorkList_PHILL" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show vbModal

End Sub



Function GetOrderSeqCode(argExamDt As String, argPID As String, argPCD As String) As String
    Dim RS As ADODB.Recordset
    
    '-- SEQ 가져오기
    
          SQL = "SELECT /*+ INDEX(rslt scrrslth_ux1) INDEX (coif scccoifm_ix1) */" & vbCr
    SQL = SQL & "       rslt.smp_no, rslt.prcp_seq, rslt.exam_seq, rslt.rept_seq, rslt.cd, rslt.pt_no, rslt.exam_stus, rslt.mach_rslt, rslt.exam_rslt ," & vbCr
    SQL = SQL & "       coif.exam_nm, prex.acp_dt, ptbs.pt_nm, ptbs.ssn_1, ptbs.ssn_2, xpsl.pt_no, " & vbCr
    SQL = SQL & "       DECODE(xpsl.gnl_add_typ_cd,'3','I',xpsl.prcp_knd_cd), xpsl.adms_ymd, xpsl.mn_sub_typ_cd, xpsl.med_dpt_cd, xpsl.med_ymd, coif.spc_cd, codm.cd_desc" & vbCr
    SQL = SQL & "  FROM scrrslth rslt, scccoifm coif, scccodem codm, scrprexh prex, mosxpslh xpsl, pmcptbsm ptbs" & vbCr
    SQL = SQL & " WHERE rslt.hos_org_no   = '" & gHOSP.HOSPCD & "'" & vbCr & vbCr
    SQL = SQL & "  AND SUBSTR(prex.acp_dt,1,8) BETWEEN '" & argExamDt & "' AND '" & argExamDt & "'" & vbCr
    SQL = SQL & "  AND rslt.smp_no       = '" & argPID & "'" & vbCr
    SQL = SQL & "  AND rslt.cd           = '" & argPCD & "'" & vbCr
    SQL = SQL & "  AND rslt.exam_stus  IN ('0','1','2')" & vbCr
    SQL = SQL & "  AND coif.hos_org_no   = rslt.hos_org_no" & vbCr
    SQL = SQL & "  AND coif.exam_cd      = rslt.cd" & vbCr
    SQL = SQL & "  AND SUBSTR(prex.acp_dt,1,8) BETWEEN coif.fr_dt AND coif.to_dt" & vbCr
    SQL = SQL & "  AND coif.exam_mach_cd = '" & gHOSP.MACHCD & "'" & vbCr
    SQL = SQL & "  AND codm.hos_org_no   = coif.hos_org_no" & vbCr
    SQL = SQL & "  AND codm.typ_cd       = '02'" & vbCr
    SQL = SQL & "  AND codm.cd           = coif.spc_cd" & vbCr
    SQL = SQL & "  AND SUBSTR(prex.acp_dt,1,8) BETWEEN codm.fr_dt AND codm.to_dt" & vbCr
    SQL = SQL & "  AND prex.hos_org_no   = rslt.hos_org_no" & vbCr
    SQL = SQL & "  AND prex.smp_no       = rslt.smp_no" & vbCr
    SQL = SQL & "  AND prex.prcp_seq     = rslt.prcp_seq" & vbCr
    SQL = SQL & "  AND prex.exam_seq     = rslt.exam_seq" & vbCr
    SQL = SQL & "  AND xpsl.hos_org_no   = prex.hos_org_no" & vbCr
    SQL = SQL & "  AND xpsl.smp_no       = prex.smp_no" & vbCr
    SQL = SQL & "  AND xpsl.acp_no       = prex.prcp_seq" & vbCr
    SQL = SQL & "  AND xpsl.prcp_typ_cd IN ('O','C')" & vbCr
    SQL = SQL & "  AND ptbs.hos_org_no   = prex.hos_org_no" & vbCr
    SQL = SQL & "  AND ptbs.pt_no        = prex.pt_no" & vbCr

    Call SetSQLData("SEQ찾기", SQL)

    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            GetOrderSeqCode = GetOrderSeqCode & Trim(RS.Fields("prcp_seq")) & "|" & Trim(RS.Fields("exam_seq")) & "|" & Trim(RS.Fields("rept_seq")) & "|"
            RS.MoveNext
        Loop
    End If
    
    If GetOrderSeqCode <> "" Then
        GetOrderSeqCode = Mid(GetOrderSeqCode, 1, Len(GetOrderSeqCode) - 1)
    End If
    
    Set RS = Nothing
    
End Function

Public Function getEASYSJudge(ByVal pOrdCD As String, ByVal pResult As String) As String
    Dim RSJ         As ADODB.Recordset
    Dim strLow      As String
    Dim strHigh     As String
    
    getEASYSJudge = ""
    
          SQL = "Select REFLOW, REFHIGH  "
    SQL = SQL & "  From EQPMASTER"
    SQL = SQL & " Where EQUIPCD = '" & gHOSP.MACHCD & "' "
    SQL = SQL & "   And TESTCODE =  '" & pOrdCD & "'"
    
    Set RSJ = New ADODB.Recordset
    Set RSJ = AdoCn_Local.Execute(SQL, , 1)
    If Not RSJ.EOF = True And Not RSJ.BOF = True Then
        strLow = Trim(RSJ.Fields("REFLOW") & "")
        strHigh = Trim(RSJ.Fields("REFHIGH") & "")
        
        If strLow <> "" And strHigh <> "" And pResult <> "" And IsNumeric(strLow) And IsNumeric(strHigh) And IsNumeric(pResult) Then
            If Val(pResult) > Val(strHigh) Then
                getEASYSJudge = "H"
            ElseIf Val(pResult) < Val(strLow) Then
                getEASYSJudge = "L"
            Else
                getEASYSJudge = " "
            End If
        Else
            getEASYSJudge = " "
        End If
    Else
        getEASYSJudge = ""
    End If
        
    RSJ.Close
    
End Function


Public Sub GetWorkList_MSINFOTEC(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As Object)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean

    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strTestCds  As String
    
On Error GoTo ErrHandle

    Screen.MousePointer = 11
    blnSame = False
    strTestCds = ""

    SQL = ""
    SQL = SQL & "SELECT R.ORDT          AS HOSPDATE                     " & vbCrLf
    SQL = SQL & "     , R.SPNO          AS BARCODE                      " & vbCrLf
    SQL = SQL & "     , R.PAID          AS PID                          " & vbCrLf
    SQL = SQL & "     , R.NWNO          AS CHARTNO                      " & vbCrLf
    SQL = SQL & "     , P.PANM          AS PNAME                        " & vbCrLf
    SQL = SQL & "     , P.SEXS          AS SEX                          " & vbCrLf
    SQL = SQL & "     , P.AGES          AS AGE                          " & vbCrLf
    SQL = SQL & "     , COUNT(R.ORCD)   AS CNT                          " & vbCrLf
    SQL = SQL & "  FROM emr.LRESULT R                                   " & vbCrLf
    SQL = SQL & "     , emr.APATINF P                                   " & vbCrLf
    SQL = SQL & " WHERE R.ORDT BETWEEN '" & pFrom & "' AND '" & pTo & "'" & vbCrLf
    SQL = SQL & "   AND R.PAID  = P.PAID                                " & vbCrLf
    SQL = SQL & "   AND R.OKFL  <> 'Y'                                  " & vbCrLf   '-- 결과확정유무 (Y / N)
    SQL = SQL & "   AND R.ORCD  IN (" & gAllTestCd & ")                 " & vbCrLf
    SQL = SQL & "   AND (R.RSFL IS NULL OR R.RSFL = 'N' OR R.RSFL = '') " & vbCrLf
    SQL = SQL & " GROUP BY R.ORDT                                       " & vbCrLf
    SQL = SQL & "        , R.SPNO                                       " & vbCrLf
    SQL = SQL & "        , R.PAID                                       " & vbCrLf
    SQL = SQL & "        , R.NWNO                                       " & vbCrLf
    SQL = SQL & "        , P.PANM                                       " & vbCrLf
    SQL = SQL & "        , P.SEXS                                       " & vbCrLf
    SQL = SQL & "        , P.AGES                                       " & vbCrLf
    SQL = SQL & " ORDER BY R.ORDT, R.PAID, P.PANM                       " & vbCrLf

    Call SetSQLData("워크조회", SQL, "")

    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then

        SPD.MaxRows = 0

        Do Until RS.EOF
            With SPD
                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strBarcode = GetText(SPD, i, colBARCODE)
                    If Trim(RS("HOSPDATE")) = strHospDate And Trim(RS("BARCODE")) = strBarcode Then
                        blnSame = True
                    End If
                Next

                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows

                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                    SetText SPD, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    
                End If
                
            End With

            blnSame = False

            DoEvents

            RS.MoveNext
        Loop
    Else
        frmInterface.lblComStatus.Caption = "워크리스트 조회 대상자가 없습니다."
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetWorkList_MSINFOTEC" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show vbModal

End Sub

Public Sub GetWorkList_EONM(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As Object)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean

    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strTestCds  As String
    
On Error GoTo ErrHandle

    Screen.MousePointer = 11
    blnSame = False
    strTestCds = ""

'    SQL = SQL & "   AND O.H141_NOTYYN  = 'N'                                    " & vbCrLf
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       O.H141_ODRDAT           AS HOSPDATE                     " & vbCrLf
    SQL = SQL & "      ,O.H141_TSAMPLENO        AS BARCODE                      " & vbCrLf
    SQL = SQL & "      ,O.H141_SEQNO            AS PID                          " & vbCrLf
    SQL = SQL & "      ,P.A110_CHARTNO          AS CHARTNO                      " & vbCrLf
    SQL = SQL & "      ,P.A110_PATNM            AS PNAME                        " & vbCrLf
    SQL = SQL & "      ,P.A110_JUMIN1           AS AGE                          " & vbCrLf
    SQL = SQL & "      ,P.A110_SEX              AS SEX                          " & vbCrLf
    SQL = SQL & "      ,COUNT(O.H141_SUGACD)    AS CNT                          " & vbCrLf
    SQL = SQL & "  FROM TB_H141_LISTAKEBODY O                                   " & vbCrLf
    SQL = SQL & "     , TB_A110_PATINFO     P                                   " & vbCrLf
    SQL = SQL & " Where O.H141_ODRDAT BETWEEN '" & pFrom & "' AND '" & pTo & "' " & vbCrLf
    SQL = SQL & "   AND P.A110_CHARTNO  = O.H141_CHARTNO                        " & vbCrLf
    SQL = SQL & "   AND O.H141_NOTYYN   IN ('N','T')                            " & vbCrLf '결과대기:T
    SQL = SQL & "   And O.H141_SUGACD   IN (" & gAllTestCd & ")                 " & vbCrLf
    SQL = SQL & " Group By O.H141_ODRDAT                                        " & vbCrLf
    SQL = SQL & "        , O.H141_TSAMPLENO                                     " & vbCrLf
    SQL = SQL & "        , O.H141_SEQNO                                         " & vbCrLf
    SQL = SQL & "        , P.A110_CHARTNO                                       " & vbCrLf
    SQL = SQL & "        , P.A110_PATNM                                         " & vbCrLf
    SQL = SQL & "        , P.A110_JUMIN1                                        " & vbCrLf
    SQL = SQL & "        , P.A110_SEX                                           " & vbCrLf
    SQL = SQL & " Order By O.H141_ODRDAT, O.H141_SEQNO                          " & vbCrLf

    Call SetSQLData("워크조회", SQL, "")

    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then

        SPD.MaxRows = 0

        Do Until RS.EOF
            With SPD
                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strBarcode = GetText(SPD, i, colBARCODE)
                    If Trim(RS("HOSPDATE")) = strHospDate And Trim(RS("BARCODE")) = strBarcode Then
                        blnSame = True
                    End If
                Next

                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows

                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                    SetText SPD, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    
                End If
                
            End With

            blnSame = False

            DoEvents

            RS.MoveNext
        Loop
    Else
        frmInterface.lblComStatus.Caption = "워크리스트 조회 대상자가 없습니다."
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetWorkList_EONM" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    
    frmErrMsg.Show vbModal

End Sub

Public Sub GetWorkList_AMIS(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As Object)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean

    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strTestCds  As String
    
On Error GoTo ErrHandle

    Screen.MousePointer = 11
    blnSame = False
    strTestCds = ""

    'SQL = SQL & "   AND R.ORDERCODE      IN (" & gAllOrdCd & ")             " & vbCrLf
'    SQL = SQL & "   AND R.RESULTFLAG    = 0                                 " & vbCrLf
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT"
    SQL = SQL & "       O.ACPTDATE              AS HOSPDATE                 " & vbCrLf
    SQL = SQL & "     , R.SPCMNO                AS BARCODE                  " & vbCrLf
    SQL = SQL & "     , P.PATID                 AS PID                      " & vbCrLf
    SQL = SQL & "     , P.PATNAME               AS PNAME                    " & vbCrLf
    SQL = SQL & "     , P.SEX                   AS SEX                      " & vbCrLf
    SQL = SQL & "     , COUNT(R.RESULTITEMCODE) AS CNT                      " & vbCrLf
    SQL = SQL & "  FROM REGISTINFOS O                                       " & vbCrLf
    SQL = SQL & "     , RESULTOFNUM R                                       " & vbCrLf
    SQL = SQL & "     , PATMST      P                                       " & vbCrLf
    SQL = SQL & " WHERE O.ACPTDATE BETWEEN '" & pFrom & "' and '" & pTo & "'" & vbCrLf
    SQL = SQL & "   AND O.ACPTDATE  = R.ACPTDATE                            " & vbCrLf
    SQL = SQL & "   AND O.PATID     = R.PATID                               " & vbCrLf
    SQL = SQL & "   AND O.ACPTSEQ   = R.ACPTSEQ                             " & vbCrLf
    SQL = SQL & "   AND O.PATID     = P.PATID                               " & vbCrLf
    SQL = SQL & "   AND O.CLAS          = 4                                 " & vbCrLf '임상병리
    SQL = SQL & "   AND R.RESULTITEMCODE IN (" & gAllTestCd & ")            " & vbCrLf
    SQL = SQL & "   AND (R.NUMRESULTVAL = '' OR R.NUMRESULTVAL IS NULL)     " & vbCrLf
    SQL = SQL & " GROUP BY O.ACPTDATE, R.SPCMNO, P.PATID, P.PATNAME, P.SEX  " & vbCrLf
    SQL = SQL & " ORDER BY O.ACPTDATE, R.SPCMNO                             " & vbCrLf

    Call SetSQLData("워크조회", SQL, "")

    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then

        SPD.MaxRows = 0

        Do Until RS.EOF
            With SPD
                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strBarcode = GetText(SPD, i, colBARCODE)
                    If Trim(RS("HOSPDATE")) = strHospDate And Trim(RS("BARCODE")) = strBarcode Then
                        blnSame = True
                    End If
                Next

                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows

                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    
                End If
                
            End With

            blnSame = False

            DoEvents

            RS.MoveNext
        Loop
    Else
        frmInterface.lblComStatus.Caption = "워크리스트 조회 대상자가 없습니다."
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetWorkList_AMIS" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    
    frmErrMsg.Show vbModal

End Sub

Public Sub GetWorkList_JWINFO(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As Object)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean

    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strTestCds  As String
    
On Error GoTo ErrHandle

    Screen.MousePointer = 11
    blnSame = False
    strTestCds = ""
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       a.RECEIPTDATE           AS HOSPDATE " & vbCrLf
    SQL = SQL & "     , a.SPECIMENNUM           AS BARCODE  " & vbCrLf
    SQL = SQL & "     , a.RECEIPTNO             AS CHARTNO  " & vbCrLf
    SQL = SQL & "     , a.PTNO                  AS PID      " & vbCrLf
    SQL = SQL & "     , a.SNAME                 AS PNAME    " & vbCrLf
    SQL = SQL & "     , COUNT(a.LABCODE)        AS CNT      " & vbCrLf
    SQL = SQL & "   FROM SLA_LabMaster a, SLA_LabResult b   " & vbCrLf
    SQL = SQL & " WHERE a.RECEIPTNO     = b.RECEIPTNO       " & vbCrLf
    SQL = SQL & "   AND a.ORDERCODE     = b.ORDERCODE       " & vbCrLf
    SQL = SQL & "   and a.RECEIPTDATE   = b.RECEIPTDATE     " & vbCrLf
    SQL = SQL & "   AND a.SPECIMENNUM   = b.SPECIMENNUM     " & vbCrLf
    SQL = SQL & "   AND a.RECEIPTDATE BETWEEN '" & Format(pFrom, "####-##-##") & "' and '" & Format(pTo, "####-##-##") & "'" & vbCrLf
    SQL = SQL & "   AND b.LABCODE IN (" & gAllTestCd & ")   " & vbCrLf
    SQL = SQL & "   AND a.JSTATUS < '3'                     " & vbCrLf
    SQL = SQL & " GROUP BY a.RECEIPTDATE, a.SPECIMENNUM, a.RECEIPTNO, a.IPDOPD, a.PTNO, a.SNAME " & vbCrLf
    SQL = SQL & " ORDER BY a.RECEIPTDATE,a.SPECIMENNUM "
    
    Call SetSQLData("워크조회", SQL, "")

    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then

        SPD.MaxRows = 0

        Do Until RS.EOF
            With SPD
                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strBarcode = GetText(SPD, i, colBARCODE)
                    If Trim(RS("HOSPDATE")) = strHospDate And Trim(RS("BARCODE")) = strBarcode Then
                        blnSame = True
                    End If
                Next

                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows

                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    
                End If
                
            End With

            blnSame = False

            DoEvents

            RS.MoveNext
        Loop
    Else
        frmInterface.lblComStatus.Caption = "워크리스트 조회 대상자가 없습니다."
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetWorkList_JWINFO" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    
    frmErrMsg.Show vbModal

End Sub

Public Sub GetWorkList_MEDICHART(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As Object)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean

    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strChartNo  As String
    Dim strTestCds  As String
    
On Error GoTo ErrHandle

    Screen.MousePointer = 11
    blnSame = False
    strTestCds = ""

'    SQL = SQL & "     , c.진료상태                          AS STATE        " & vbCrLf

    SQL = ""
    SQL = SQL & "Select DISTINCT "
    SQL = SQL & "       (a.진료년 + a.진료월 + a.진료일)    AS HOSPDATE     " & vbCrLf
    SQL = SQL & "     , a.챠트번호                          AS CHARTNO      " & vbCrLf
    SQL = SQL & "     , b.수진자명                          AS PNAME        " & vbCrLf
    SQL = SQL & "     , b.주민등록번호                      AS PJUMIN       " & vbCrLf
    SQL = SQL & "     , COUNT(a.처방코드)                   AS CNT          " & vbCrLf
    SQL = SQL & "  From TB_검사항목 a                                       " & vbCrLf
    SQL = SQL & "     , TB_인적사항 b                                       " & vbCrLf
    SQL = SQL & "     , TB_진료기본 c                                       " & vbCrLf
    SQL = SQL & " Where (a.진료년 + a.진료월 + a.진료일) >= '" & pFrom & "' " & vbCrLf
    SQL = SQL & "   And (a.진료년 + a.진료월 + a.진료일) <= '" & pTo & "'   " & vbCrLf
    SQL = SQL & "   And a.처방번호 > 0                                      " & vbCrLf
    SQL = SQL & "   And c.진료상태 IN ('1','5','6','7','8','9')             " & vbCrLf
    'SQL = SQL & "   And (a.처방코드 + a.서브코드) IN (" & gAllTestCd & ")   " & vbCrLf
    SQL = SQL & "   And (a.처방코드 + '|' + a.서브코드) IN (" & gAllTestCd & ")   " & vbCrLf
    SQL = SQL & "   And (a.검사결과 IS NULL OR a.검사결과 = '')             " & vbCrLf
    SQL = SQL & "   And a.진료년    = c.진료년                              " & vbCrLf
    SQL = SQL & "   And a.진료월    = c.진료월                              " & vbCrLf
    SQL = SQL & "   And a.진료일    = c.진료일                              " & vbCrLf
    SQL = SQL & "   And a.챠트번호  = c.챠트번호                            " & vbCrLf
    SQL = SQL & "   And a.챠트번호  = b.챠트번호                            " & vbCrLf
    SQL = SQL & "   And (a.검사결과 IS NULL OR a.검사결과 = '')             " & vbCrLf
    SQL = SQL & " GROUP BY HOSPDATE, a.챠트번호, b.수진자명, b.주민등록번호 " & vbCrLf   ', c.진료상태
    SQL = SQL & " Order By a.진료년, a.진료월, a.진료일, b.수진자명         " & vbCrLf

    Call SetSQLData("워크조회", SQL, "")

    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then

        SPD.MaxRows = 0

        Do Until RS.EOF
            With SPD
                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strBarcode = GetText(SPD, i, colBARCODE)
                    strChartNo = GetText(SPD, i, colCHARTNO)
                    If Trim(RS("HOSPDATE")) = strHospDate And Trim(RS("CHARTNO")) = strChartNo Then
                        blnSame = True
                    End If
                Next

                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows

                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    'SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colPID
                    'SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("PJUMIN")) & "", intRow, colPAGE
                    'SetText SPD, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    
                End If
                
            End With

            blnSame = False

            DoEvents

            RS.MoveNext
        Loop
    Else
        frmInterface.lblComStatus.Caption = "워크리스트 조회 대상자가 없습니다."
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetWorkList_MEDICHART" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    
    frmErrMsg.Show vbModal

End Sub


Public Sub GetWorkList_KCHART(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean

    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strTestCds  As String
    Dim strItems    As String
    
On Error GoTo ErrHandle

    Screen.MousePointer = 11
    blnSame = False
    strTestCds = ""
    
    pFrom = Format(pFrom, "####-##-##")
    pTo = Format(pTo, "####-##-##")

    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       J.접수일자          AS HOSPDATE                         " & vbCrLf
    SQL = SQL & "     , L.검체번호          AS BARCODE                          " & vbCrLf
    SQL = SQL & "     , A.챠트번호          AS CHARTNO                          " & vbCrLf
    SQL = SQL & "     , J.접수번호          AS PID                              " & vbCrLf
    SQL = SQL & "     , A.환자이름          AS PNAME                            " & vbCrLf
    SQL = SQL & "     , A.환자성별          AS SEX                              " & vbCrLf
    SQL = SQL & "     , A.환자나이          AS AGE                              " & vbCrLf
    SQL = SQL & "     , COUNT(L.처방코드)   AS CNT                              " & vbCrLf
    SQL = SQL & "  FROM         TB_진료검사 L                                   " & vbCrLf
    SQL = SQL & "   INNER JOIN  TB_진료지원 J ON (L.진료지원ID = J.진료지원ID)  " & vbCrLf
    SQL = SQL & "   INNER JOIN  TB_진료일반 A ON (J.진료일자   = A.진료일자     " & vbCrLf
    SQL = SQL & "                            AND  J.챠트번호   = A.챠트번호     " & vbCrLf
    SQL = SQL & "                            AND  J.진료번호   = A.진료번호)    " & vbCrLf
    SQL = SQL & " Where J.접수일자 BETWEEN '" & pFrom & "' and '" & pTo & "'    " & vbCrLf
    SQL = SQL & "   AND L.검사상태 < 5                                          " & vbCrLf
    SQL = SQL & "   AND L.처방코드 + L.서브코드 IN (" & gAllTestCd & ")         " & vbCrLf
    SQL = SQL & " GROUP BY J.접수일자, L.검체번호, A.챠트번호, J.접수번호, A.환자이름, A.환자성별, A.환자나이 " & vbCrLf
'    SQL = SQL & " ORDER BY J.접수일자, L.검체번호                               " & vbCrLf

    Call SetSQLData("워크조회", SQL, "")

    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then

        SPD.MaxRows = 0

        Do Until RS.EOF
            With SPD
                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strBarcode = GetText(SPD, i, colBARCODE)
                    If Trim(RS("HOSPDATE")) = strHospDate And Trim(RS("BARCODE")) = strBarcode Then
                        blnSame = True
                    End If
                Next

                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows

                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                    SetText SPD, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    
                    '장비에서 오더요청이 안오는 배치오더용
                    Select Case gHOSP.MACHNM
                        Case "ACCESS2"
                            strItems = GetEquipExamCode_ACCESS2(gHOSP.MACHCD, "")
                            'Call SetText(SPD, strItems, intRow, colDEPT)
                            Call SetTag(SPD, strItems, intRow, colSTATE)
                        
                        Case "PPC300N"
                            strItems = GetEquipExamCode_PPC300N(gHOSP.MACHCD, "")
                            Call SetTag(SPD, strItems, intRow, colSTATE)
                            'Call SetText(SPD, strItems, intRow, colDEPT)
                    End Select
                    
                End If
                
            End With

            blnSame = False

            DoEvents

            RS.MoveNext
        Loop
    Else
        frmInterface.lblComStatus.Caption = "워크리스트 조회 대상자가 없습니다."
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetWorkList_KCHART" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    
    frmErrMsg.Show vbModal

End Sub

Public Sub GetWorkList_SY(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean

    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strTestCds  As String
    Dim strItems    As String
    
On Error GoTo ErrHandle

    Screen.MousePointer = 11
    blnSame = False
    strTestCds = ""
   

    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       CONVERT(NVARCHAR(50),M.접수일자,112) AS HOSPDATE"
    SQL = SQL & "     , M.접수번호 AS PID"
    'SQL = SQL & "     ,'' AS 의뢰처"
    SQL = SQL & "     , M.차트번호 AS CHARTNO"
    SQL = SQL & "     , M.성명 AS PNAME"
    SQL = SQL & "     , M.성별 AS SEX"
    SQL = SQL & "     , M.나이 AS AGE"
    SQL = SQL & "     , E.검사코드 AS ITEM " & vbCr
    SQL = SQL & "  FROM VW_검사접수 M, VW_검사결과 R, VW_검사코드 E  " & vbCr
    SQL = SQL & " WHERE M.접수일자 = CONVERT(DATE, '" & pFrom & "')" & vbCr
    SQL = SQL & "   AND E.학부코드 = '" & gHOSP.PARTCD & "' " & vbCr
    SQL = SQL & "   AND E.검사코드 IN (" & gAllTestCd & ") " & vbCr
    SQL = SQL & "   AND isnull(R.보고여부, 'N') <> 'Y' " & vbCr
    SQL = SQL & "   AND (R.결과값 is null or R.결과값 = '') " & vbCr
    SQL = SQL & "   AND M.접수일자 = R.접수일자 " & vbCr
    SQL = SQL & "   AND M.접수번호 = R.접수번호 " & vbCr
    SQL = SQL & " ORDER BY HOSPDATE, PID "
    
    Call SetSQLData("워크조회", SQL, "")

    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then

        SPD.MaxRows = 0

        Do Until RS.EOF
            With SPD
                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strBarcode = GetText(SPD, i, colBARCODE)
                    If Trim(RS("HOSPDATE")) = strHospDate And Mid(Trim(RS.Fields("HOSPDATE")), 1, 8) & PedLeftStr(Trim(RS.Fields("PID")), 5, "0") = strBarcode Then
                        blnSame = True
                        strItems = strItems & GetTestNm(Trim(RS.Fields("ITEM"))) & "/"
                        iCnt = iCnt + 1
                    End If
                Next

                If blnSame = False Then
                    If strItems <> "" Then
                        SetText SPD, strItems, intRow, colITEMS
                        SetText SPD, CStr(iCnt), intRow, colOCNT
                    End If
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows
                    strItems = ""
                    iCnt = 0
                    
                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & PedLeftStr(Trim(RS.Fields("PID")), 5, "0"), intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                    SetText SPD, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
                    
                    strItems = strItems & GetTestNm(Trim(RS.Fields("ITEM"))) & "/"
                    iCnt = iCnt + 1
                End If
                
            End With

            blnSame = False

            DoEvents

            RS.MoveNext
        Loop
    Else
        frmInterface.lblComStatus.Caption = "워크리스트 조회 대상자가 없습니다."
    End If

    RS.Close

    If strItems <> "" Then
        SetText SPD, strItems, intRow, colITEMS
        SetText SPD, CStr(iCnt), intRow, colOCNT
    End If
    
    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetWorkList_KCHART" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    
    frmErrMsg.Show vbModal

End Sub


    
'접수일시(FROM)     String      필수
'접수일시(TO)       String      필수
'슬립코드           String      필수
'작업번호(FROM)     String      필수
'작업번호(TO)       String      필수
'검사코드       String      필수아님, 복수시 '|' 로구분

'{"rcpnDvsn":"C","slipCd":"L20","spcmStat":"2","sex":"M","viewRslt":"","pid":"08015508","ward":"","rcpnDt":"2020-07-15 11:08:44","workNo":"20200715L2000I0159","spcmNm":"CBC/EDTA tube","exmnCd":"LA101827","rexmYn":"","dobr":"19831116","rsltStat":"","ptNm":"정원상","medDp":"FM","brcdLablNo":"2024406394","spcmSqno":"","age":"36","spcmCd":"002"},
'{
' "result": [{
'  "rcpnDtf" : "20180310"
'  "brcdLablNo" : "1820028641"
'  "workNo" : "20180323L2000I0001"
'  "slipCd" : "L20"
'  "pid" : "00012556"
'  "ptNm" : "우동*"
'  "dobr" : "19770529"
'  "sex" : "M"
'  "age" : "40"
'  "rcpnDvsn" : "I"
'  "medDp" : "(임시)심장내과"
'  "ward" : "240병동/01"
'  "rcpnDt" : "2018-03-23 14:28:08"
'  "spcmCd" : "034"
'  "spcmNm" : "EDTA W.B"
'  "exmnCd" : "GL2002"
'  "viewRslt" : "0"
'  "rexmYn" : "0"
'  "spcmStat" : "2"
'  "rsltStat" : "0"
' }],
' "serviceCode": 100,
' "Count": 1,
' "serviceMsg": ""
'}
Public Sub GetWorkList_BITJSON(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As Object, Optional ByVal pFromNo As String = "", Optional ByVal pToNo As String = "")
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean

    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strItems    As String
    
    Dim strParam(5) As Variant
    Dim strReturn   As Variant
    Dim strJData    As Variant
    
    Dim strDate     As String
    Dim strWN       As String
    Dim strPtNm     As String
    Dim strSex      As String
    Dim strAge      As String
    Dim strPID      As String
    Dim strDept     As String
    Dim strExmnCd   As String
    
On Error GoTo ErrHandle

    Screen.MousePointer = 11
    blnSame = False

    strItems = Replace(gAllTestCd, "'", "")
    strItems = Replace(strItems, ",", "|")
    
    strParam(0) = pFrom                                     '20180310
    strParam(1) = pTo                                       '20180330
    strParam(2) = gHOSP.PARTCD                              'L20
    strParam(3) = pFrom & gHOSP.PARTCD & "00I" & pFromNo     '20180310L2000I0001
    strParam(4) = pTo & gHOSP.PARTCD & "00I" & pToNo       '20180330L2000I0999
    strParam(5) = strItems                                  'GL2002|L2011|L2018
    Call SetSQLData("워크조회", strParam(0) & "," & strParam(1) & "," & strParam(2) & "," & strParam(3) & "," & strParam(4) & "," & strParam(5), "A")
    strReturn = JsonSend("WORKLIST", strParam)
    
    Call SetSQLData("워크조회", strReturn, "A")
    Call getJsonVar(CStr(strReturn))
    
    strJData = Split(mJsonData, vbCr)
    
    SPD.MaxRows = 0
            
            
            
    If strReturn = "" Then
        '조회대상자 없음
        frmInterface.lblComStatus.Caption = "워크리스트 조회 대상자가 없습니다."
    Else
        strReturn = Split(strReturn, vbCr)
    
        If mJsonData <> "" Then
            With SPD
                .ReDraw = False
                For i = 0 To UBound(strJData)
                    Debug.Print Trim(mGetP(strJData(i), 1, "@"))
                    Select Case Trim(mGetP(strJData(i), 1, "@"))
                        Case "rcpnDt":      strDate = Trim(mGetP(strJData(i), 2, "@"))
                        Case "workNo":      strWN = Trim(mGetP(strJData(i), 2, "@"))
                        Case "ptNm":        strPtNm = Trim(mGetP(strJData(i), 2, "@"))
                        Case "brcdLablNo":  strBarcode = Trim(mGetP(strJData(i), 2, "@"))
                        Case "sex":         strSex = Trim(mGetP(strJData(i), 2, "@"))
                        Case "age":         strAge = Trim(mGetP(strJData(i), 2, "@"))
                        Case "pid":         strPID = Trim(mGetP(strJData(i), 2, "@"))
                        Case "medDp":       strDept = Trim(mGetP(strJData(i), 2, "@"))
                        Case "exmnCd":      strExmnCd = Trim(mGetP(strJData(i), 2, "@"))
                        Case "spcmCd": '마지막
                            blnSame = False
                            For iCnt = 0 To SPD.DataRowCnt
                                If Mid(strDate, 1, 10) = GetText(SPD, iCnt, colHOSPDATE) And strBarcode = GetText(SPD, iCnt, colBARCODE) Then
                                    blnSame = True
                                    strItems = strItems & GetTestNm(strExmnCd) & "/"
                                    Exit For
                                End If
                            Next
                            
                            If blnSame = False Then
                                If strItems <> "" Then
                                    SetText SPD, strItems, intRow, colITEMS
                                End If
                                
                                .MaxRows = .MaxRows + 1
                                intRow = .MaxRows
                                strItems = ""
                                
                                SetText SPD, "1", intRow, colCHECKBOX
                                SetText SPD, Mid(strDate, 1, 10), intRow, colHOSPDATE
                                SetText SPD, strBarcode, intRow, colBARCODE
                                SetText SPD, strWN, intRow, colCHARTNO
                                SetText SPD, strPID, intRow, colPID
                                SetText SPD, strPtNm, intRow, colPNAME
                                SetText SPD, strSex, intRow, colPSEX
                                SetText SPD, strAge, intRow, colPAGE
                                'SetText SPD, GetSampleITEM(intRow), intRow, colITEMS
                                strItems = strItems & GetTestNm(strExmnCd) & "/"
                            End If
                    End Select
                Next
            End With
        End If
    End If

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetWorkList_BITJSON" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    
    frmErrMsg.Show vbModal

End Sub


'    send = ""
'    send = send & "20180116|20180116|I0101|201801160027|WMD0910173|10117020|일억수산수족관수|I010100000051554|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160027|WMD0910174|10117020|일억수산수족관수|I010100000051554|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160027|WMD0910175|10117020|일억수산수족관수|I010100000051554|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160027|WMD0910176|10117020|일억수산수족관수|I010100000051554|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160027|WMD0910181|10117020|일억수산수족관수|I010100000051554|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160027|WMD0910182|10117020|일억수산수족관수|I010100000051554|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160027|WMD0910183|10117020|일억수산수족관수|I010100000051554|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160027|WMD0910184|10117020|일억수산수족관수|I010100000051554|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160027|WMD0910188|10117020|일억수산수족관수|I010100000051554|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160028|WMD0910173|10117021|삼양닭집양념소스|I010100000051555|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160028|WMD0910174|10117021|삼양닭집양념소스|I010100000051555|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160028|WMD0910175|10117021|삼양닭집양념소스|I010100000051555|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160028|WMD0910176|10117021|삼양닭집양념소스|I010100000051555|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160028|WMD0910181|10117021|삼양닭집양념소스|I010100000051555|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160028|WMD0910182|10117021|삼양닭집양념소스|I010100000051555|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160028|WMD0910183|10117021|삼양닭집양념소스|I010100000051555|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160028|WMD0910184|10117021|삼양닭집양념소스|I010100000051555|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180116|20180116|I0101|201801160028|WMD0910188|10117021|삼양닭집양념소스|I010100000051555|000000|0AAEBA1DMtEZAeg08JGGXSSgZ|0|S2|"
'    send = send & "20180129|20180129|I0101|201801290036|WMD0910185|00216951|송혜경|2008101671822354|861121|2AAEjAFRdDJ/dcYvj8o8rt9ND|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290036|WMD0910186|00216951|송혜경|2008101671822354|861121|2AAEjAFRdDJ/dcYvj8o8rt9ND|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290036|WMD0910187|00216951|송혜경|2008101671822354|861121|2AAEjAFRdDJ/dcYvj8o8rt9ND|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290037|WMD0910185|10118013|이승재|I010100000052327|841019|1AAEDWTDOLpBIjgwhDxQL5zX2|1|S2|"
'    send = send & "20180129|20180129|I0101|201801290037|WMD0910186|10118013|이승재|I010100000052327|841019|1AAEDWTDOLpBIjgwhDxQL5zX2|1|S2|"
'    send = send & "20180129|20180129|I0101|201801290037|WMD0910187|10118013|이승재|I010100000052327|841019|1AAEDWTDOLpBIjgwhDxQL5zX2|1|S2|"
'    send = send & "20180129|20180129|I0101|201801290084|WMD0910185|10118018|전소희|2008091615664689|930710|2AAEkiHJC73gU9QjNUgoqzxmA|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290084|WMD0910186|10118018|전소희|2008091615664689|930710|2AAEkiHJC73gU9QjNUgoqzxmA|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290084|WMD0910187|10118018|전소희|2008091615664689|930710|2AAEkiHJC73gU9QjNUgoqzxmA|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290085|WMD0910185|10118017|윤찬미|I130100000375849|920629|2AAEGhRZybQ/G5kOev+0BjH+j|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290085|WMD0910186|10118017|윤찬미|I130100000375849|920629|2AAEGhRZybQ/G5kOev+0BjH+j|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290085|WMD0910187|10118017|윤찬미|I130100000375849|920629|2AAEGhRZybQ/G5kOev+0BjH+j|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290103|WMD0910185|10061809|최순희|2008091015474332|770902|2AAEnACI5oeWP0AzCceZUz9vD|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290103|WMD0910186|10061809|최순희|2008091015474332|770902|2AAEnACI5oeWP0AzCceZUz9vD|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290103|WMD0910187|10061809|최순희|2008091015474332|770902|2AAEnACI5oeWP0AzCceZUz9vD|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290148|WMD0910185|00054686|조윤옥|2008061231073932|820210|2AAEfShJJOIfa9WmOBMDzQ9Zr|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290148|WMD0910186|00054686|조윤옥|2008061231073932|820210|2AAEfShJJOIfa9WmOBMDzQ9Zr|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290148|WMD0910187|00054686|조윤옥|2008061231073932|820210|2AAEfShJJOIfa9WmOBMDzQ9Zr|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290149|WMD0910185|10017135|서숙희|2008061030133445|710418|2AAHC/eTSqqBqvs5ls41xYygl|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290149|WMD0910186|10017135|서숙희|2008061030133445|710418|2AAHC/eTSqqBqvs5ls41xYygl|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290149|WMD0910187|10017135|서숙희|2008061030133445|710418|2AAHC/eTSqqBqvs5ls41xYygl|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290157|WMD0910185|00033390|박혜영|2008092918367282|890131|2AAHGfNo0OLHgnwGk9FGvVr4F|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290157|WMD0910186|00033390|박혜영|2008092918367282|890131|2AAHGfNo0OLHgnwGk9FGvVr4F|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290157|WMD0910187|00033390|박혜영|2008092918367282|890131|2AAHGfNo0OLHgnwGk9FGvVr4F|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290159|WMD0910185|10016141|김수빈|2007122700062926|880525|1AAGYRWKhoXOnk4VPINHxI7/+|1|S2|"
'    send = send & "20180129|20180129|I0101|201801290159|WMD0910186|10016141|김수빈|2007122700062926|880525|1AAGYRWKhoXOnk4VPINHxI7/+|1|S2|"
'    send = send & "20180129|20180129|I0101|201801290159|WMD0910187|10016141|김수빈|2007122700062926|880525|1AAGYRWKhoXOnk4VPINHxI7/+|1|S2|"
'    send = send & "20180129|20180129|I0101|201801290206|WMD0910185|00047737|이미연|2008092918394552|780408|2AAHpCg8s6nXHwiEHlK/KCEQE|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290206|WMD0910186|00047737|이미연|2008092918394552|780408|2AAHpCg8s6nXHwiEHlK/KCEQE|2|S2|"
'    send = send & "20180129|20180129|I0101|201801290206|WMD0910187|00047737|이미연|2008092918394552|780408|2AAHpCg8s6nXHwiEHlK/KCEQE|2|S2|"
'    strReturn = send
'Public Sub GetWorkList_HEALTH(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As Object)
'    Dim oSOAP       As MSSOAPLib30.SoapClient30
'    Dim strParam    As String
'    Dim strReturn   As Variant
'    Dim strHData    As Variant
'    Dim blnSame     As Boolean
'    Dim send
'    Dim sRet
'
'    Dim i           As Integer
'    Dim iCnt        As Integer
'    Dim intRow      As Integer
'    Dim strHospDate As String
'    Dim strBarcode  As String
'    Dim strItems    As String
'    Dim strDate     As String
'    Dim strWN       As String
'    Dim strPtNm     As String
'    Dim strSEX      As String
'    Dim strAGE      As String
'    Dim strPID      As String
'    Dim strDept     As String
'    Dim strExmnCd   As String
'
'On Error GoTo ErrHandle
'
'    Screen.MousePointer = 11
'
'    Set oSOAP = New MSSOAPLib30.SoapClient30
'    oSOAP.ClientProperty("ServerHTTPRequest") = True
'    oSOAP.MSSoapInit gHEALTH.INITURL
'
'    strParam = ""
'    strParam = strParam & "MSH|^~\&|HL7|MMS|||1||ORU^R01|1a082e2:10e59b48c04:-2cf9:27695009|P|2.3||||||8859/1" & Chr(13) & Chr(10)
'    strParam = strParam & "PID|||^" & gHOSP.MACHCD & "^" & gHOSP.USERID & "^DefaultDomain^PI" & Chr(13) & Chr(10)                   'S2     I010100037
'    strParam = strParam & "PV1||E|" & gHOSP.HOSPCD & Chr(13) & Chr(10)                                                              'I0101
'    strParam = strParam & "OBR|1||||||1" & Chr(13) & Chr(10)
'
'    strParam = Chr(11) & strParam & Chr(12) & Chr(13)
'
'    Call SetSQLData("워크조회", strParam)
'    strParam = MakeB64(strParam)
'    strReturn = oSOAP.MdbOrderList(strParam)
'    strReturn = MakeUB64(strReturn)
'    Call SetSQLData("워크조회", strReturn, "A")
'    Set oSOAP = Nothing
'    DoEvents
'
'    blnSame = False
'
'    SPD.MaxRows = 0
'
'    If strReturn = "" Then
'        '조회대상자 없음
'        frmInterface.lblComStatus.Caption = "워크리스트 조회 대상자가 없습니다."
'    Else
'        strHData = Split(strReturn, ETX)
'
'        With SPD
'            .ReDraw = False
'            For i = 0 To UBound(strHData)
'                blnSame = False
'                For iCnt = 0 To SPD.DataRowCnt
'                    If Trim(mGetP(strHData(i), 1, "|")) = GetText(SPD, iCnt, colHOSPDATE) And Trim(mGetP(strHData(i), 4, "|")) = GetText(SPD, iCnt, colBARCODE) Then
'                        blnSame = True
'                        strItems = strItems & GetTestNm(Trim(mGetP(strHData(i), 5, "|"))) & "/"
'                        Exit For
'                    End If
'                Next
'
'                If blnSame = False Then
'                    If strItems <> "" Then
'                        SetText SPD, strItems, intRow, colITEMS
'                    End If
'
'                    If Trim(mGetP(strHData(i), 1, "|")) <> "" Then
'                        .MaxRows = .MaxRows + 1
'                        intRow = .MaxRows
'                        strItems = ""
'
'                        SetText SPD, "1", intRow, colCHECKBOX
'                        SetText SPD, Trim(mGetP(strHData(i), 1, "|")), intRow, colHOSPDATE
'                        SetText SPD, Trim(mGetP(strHData(i), 4, "|")), intRow, colBARCODE
'                        SetText SPD, Trim(mGetP(strHData(i), 6, "|")), intRow, colCHARTNO
'                        SetText SPD, Trim(mGetP(strHData(i), 6, "|")), intRow, colPID
'                        SetText SPD, Trim(mGetP(strHData(i), 7, "|")), intRow, colPNAME
'                        strItems = strItems & GetTestNm(Trim(mGetP(strHData(i), 5, "|"))) & "/"
'                    End If
'                End If
'            Next
'        End With
'    End If
'
'    SPD.RowHeight(-1) = 15
'    SPD.ReDraw = True
'
'    Screen.MousePointer = 0
'
'Exit Sub
'
'ErrHandle:
'    Screen.MousePointer = 1
'
'    strErrMsg = ""
'    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetWorkList_HEALTH" & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
'    frmErrMsg.txtErr = vbNewLine & strErrMsg
'
'    frmErrMsg.Show vbModal
'
'End Sub

Private Function GetWorkList_XML() As Variant
    Dim strPath     As String
    Dim strBuffer   As String
    Dim BufChar     As String
    Dim strTmp      As String
    Dim intIdx      As Integer
    Dim i           As Long
    Dim lngBufLen   As Long
    Dim TextLine
    
On Error GoTo ErrorTrap
    
    Screen.MousePointer = 11
    blnSameRecord = False
    
    GetWorkList_XML = ""
    
    '-- 오더파일명과 경로를 지정한다.
    'strPath = "C:\UBCare\SINAI\IF\EXAMIF_IN.xml"
    strPath = gHOSP.XMLPATH & "\EXAMIF_IN.xml"
    
    Open strPath For Input As #1 ' 파일을 엽니다.
    
    Do While Not EOF(1)         ' 파일의 끝을 만날 때까지 반복합니다.
        Line Input #1, TextLine ' 변수로 데이터 행을 읽어들입니다.
        strBuffer = strBuffer & TextLine
    Loop
    
    Close #1 ' 파일을 닫습니다
 
    intIdx = 0
    lngBufLen = Len(strBuffer)
        
    For i = 1 To lngBufLen
        If intIdx = 0 Then
            BufChar = Mid$(strBuffer, i, 4)
        Else
            BufChar = Mid$(strBuffer, i + 3)
        End If
        
        If BufChar = "<검사>" Then
            intIdx = 1
            strTmp = BufChar
        Else
            strTmp = strTmp & BufChar
            If intIdx = 1 Then Exit For
        End If
    Next
    
    strTmp = Replace(strTmp, "<검사>", "")
    strTmp = Replace(strTmp, "</검사>", "|")
    
    strTmp = Replace(strTmp, "<업체>", "")
    strTmp = Replace(strTmp, "</업체>", ",")
    
    strTmp = Replace(strTmp, "<요양기관번호>", "")
    strTmp = Replace(strTmp, "</요양기관번호>", ",")
    
    strTmp = Replace(strTmp, "<차트번호>", "")
    strTmp = Replace(strTmp, "</차트번호>", ",")
    
    strTmp = Replace(strTmp, "<수진자명>", "")
    strTmp = Replace(strTmp, "</수진자명>", ",")
    
    strTmp = Replace(strTmp, "<주민등록번호>", "")
    strTmp = Replace(strTmp, "</주민등록번호>", ",")
    
    strTmp = Replace(strTmp, "<내원번호>", "")
    strTmp = Replace(strTmp, "</내원번호>", ",")
    
    strTmp = Replace(strTmp, "<의뢰일>", "")
    strTmp = Replace(strTmp, "</의뢰일>", ",")
    
    strTmp = Replace(strTmp, "<검사번호>", "")
    strTmp = Replace(strTmp, "</검사번호>", ",")
    
    strTmp = Replace(strTmp, "<검사ID>", "")
    strTmp = Replace(strTmp, "</검사ID>", ",")
    
    strTmp = Replace(strTmp, "<업체검사ID>", "")
    strTmp = Replace(strTmp, "</업체검사ID>", ",")
    
    strTmp = Replace(strTmp, "<검체>", "")
    strTmp = Replace(strTmp, "</검체>", ",")
    
    strTmp = Replace(strTmp, "<결과치>", "")
    strTmp = Replace(strTmp, "</결과치>", ",")
    
    strTmp = Replace(strTmp, "<참조치>", "")
    strTmp = Replace(strTmp, "</참조치>", ",")
    
    strTmp = Replace(strTmp, "<소견>", "")
    strTmp = Replace(strTmp, "</소견>", ",")
    
    strTmp = Replace(strTmp, "<결과일>", "")
    strTmp = Replace(strTmp, "</결과일>", ",")
    
    strTmp = Replace(strTmp, "<업체>", "")
    strTmp = Replace(strTmp, "</업체>", ",")
    
    strTmp = Replace(strTmp, "<입원외래구분>", "")
    strTmp = Replace(strTmp, "</입원외래구분>", ",")
    
    GetWorkList_XML = Split(strTmp, "|")
    
    blnSameRecord = True
    
    Screen.MousePointer = 0
    
    Exit Function
        
ErrorTrap:
    
    blnSameRecord = False
    Screen.MousePointer = 0
    
    
End Function


Public Sub GetWorkList_UBCARE(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As vaSpread)
    Dim RS          As ADODB.Recordset
    Dim RS_L        As ADODB.Recordset
    Dim RS_C        As ADODB.Recordset
    Dim blnSame     As Boolean
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strPID      As String
    Dim varXML      As Variant
    Dim varTmp      As Variant
    Dim strBarNum   As String
    Dim strJumin    As String
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    
    blnSame = False
    
    '1. XML 파일을 읽는다.
    varXML = GetWorkList_XML
    
    If blnSameRecord = True Then
        If UBound(varXML) >= 1 Then
            For i = 0 To UBound(varXML) - 1
                varTmp = Split(varXML(i), ",")
                
                '2. 해당검사코드의 채널, 검사명을 가져온다.
                SQL = ""
                SQL = SQL & "SELECT DISTINCT a.SENDCHANNEL, a.RSLTCHANNEL, a.TESTNAME   " & vbCrLf
                SQL = SQL & "  FROM EQPMASTER a, TESTMASTER b                           " & vbCrLf
                SQL = SQL & " WHERE a.RSLTCHANNEL = b.RSLTCHANNEL                       " & vbCrLf
                SQL = SQL & "   AND b.TESTCODE = '" & Trim(varTmp(8)) & "'                " & vbCrLf
                    
                Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                
                If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                    XMLInData.ComExamID = Trim(RS_L.Fields("RSLTCHANNEL").Value)
                    XMLInData.Company = varTmp(0)
                    XMLInData.HospCode = varTmp(1)
                    XMLInData.ChartNo = varTmp(2)
                    XMLInData.PatName = varTmp(3)
                    XMLInData.PatJumin = varTmp(4)
                    XMLInData.PatNo = varTmp(5)
                    XMLInData.CommDate = varTmp(6)
                    XMLInData.ExamNo = varTmp(7)
                    XMLInData.ExamID = varTmp(8)
                    'XMLInData.ComExamID = varTmp(9)
                    XMLInData.Specimen = varTmp(10)
                    XMLInData.Result = varTmp(11)
                    XMLInData.Reference = varTmp(12)
                    XMLInData.Remark = varTmp(13)
                    XMLInData.RsltDate = varTmp(14)
                    XMLInData.IOFlag = varTmp(15)
                    
                    '나이/성별 계산용
                    strJumin = LEFT(XMLInData.PatJumin, 6) & Right(XMLInData.PatJumin, 7)
                    
                    Call CalAgeSex(strJumin, Format(Date, "yyyy/mm/dd"))
                    
                    strJumin = XMLInData.PatJumin
                    
                    strBarNum = Mid(XMLInData.CommDate, 5, 4) & Format(XMLInData.ChartNo, "0000000000")

                    '-- HBA1C
                    'If XMLInData.ExamID = "C3825" Then
                    '    strBarNum = Mid(XMLInData.CommDate, 5, 4) & "H" & Format(XMLInData.ChartNo, "000000000")
                    'End If
                    
                    '3. 새로운 환자인지 확인한다.
                    SQL = ""
                    SQL = SQL & "SELECT DISTINCT CHARTNO                        " & vbCrLf
                    SQL = SQL & "  FROM UB_PATRESULT                            " & vbCrLf
                    SQL = SQL & " WHERE HOSPDATE = '" & XMLInData.CommDate & "' " & vbCrLf
                    SQL = SQL & "   AND BARCODE  = '" & strBarNum & "'          " & vbCrLf
                    SQL = SQL & "   AND CHARTNO  = '" & XMLInData.ChartNo & "'  " & vbCrLf
                    SQL = SQL & "   AND EXAMTYPE = '" & gHOSP.PARTCD & "'       " & vbCrLf
                    SQL = SQL & "   AND EXAMCODE = '" & XMLInData.ExamID & "'   " & vbCrLf
                    
                    Set RS_C = AdoCn_Local.Execute(SQL, , 1)
                    
                    If Not RS_C.EOF = True And Not RS_C.BOF = True Then
                        '4-1. 기존 환자인 경우 이름,성별,나이,검사코드 정보만 업데이트 한다.
                        '==> 처방이 변경이되거나 추가처방이 되면 <검사번호>XX</검사번호> 가 변경되므로 업데이트 해줘야 함.
                        SQL = ""
                        SQL = SQL & "Update UB_PATRESULT Set "
                        SQL = SQL & "       PNAME     = '" & XMLInData.PatName & "'  " & vbCrLf
                        SQL = SQL & "     , PSEX      = '" & mPatient.SEX & "'       " & vbCrLf
                        SQL = SQL & "     , PAGE      = '" & mPatient.AGE & "'       " & vbCrLf
                        SQL = SQL & "     , EQUIPCODE = '" & XMLInData.ComExamID & "'" & vbCrLf
                        SQL = SQL & "     , EXAMNO    = '" & XMLInData.ExamNo & "'   " & vbCrLf
                        SQL = SQL & " Where HOSPDATE  = '" & XMLInData.CommDate & "' " & vbCrLf
                        SQL = SQL & "   and BARCODE   = '" & strBarNum & "'          " & vbCrLf
                        SQL = SQL & "   and CHARTNO   = '" & XMLInData.ChartNo & "'  " & vbCrLf
                        SQL = SQL & "   and EXAMTYPE  = '" & gHOSP.PARTCD & "'       " & vbCrLf
                        SQL = SQL & "   and EXAMCODE  = '" & XMLInData.ExamID & "'   " & vbCrLf
                        'SQL = SQL & "   and EXAMNO    = '" & XMLInData.ExamNo & "'   " & vbCrLf
                    Else
                        '4-2. 새로운 환자인 경우 레코드를 만든다
                        SQL = ""
                        SQL = SQL & "INSERT INTO UB_PATRESULT (" & vbCr
                        SQL = SQL & "  SAVESEQ"                         '저장순번(날짜별)
                        SQL = SQL & ", EXAMDATE"                        '검사일자"
                        SQL = SQL & ", HOSPDATE"                        '병원접수일자"
                        SQL = SQL & ", EQUIPNO"                         '장비코드"
                        SQL = SQL & ", BARCODE              " & vbCrLf  '검체번호
                        SQL = SQL & ", EQUIPCODE"                       '검사채널"
                        SQL = SQL & ", ORDERCODE"                       '병원처방코드"
                        SQL = SQL & ", EXAMCODE"                        '병원검사코드"
                        SQL = SQL & ", EXAMSUBCODE"                     '병원검사코드(SUB)"
                        SQL = SQL & ", EXAMNAME             " & vbCrLf  '검사명
                        SQL = SQL & ", SAMPLETYPE"                      '검체유형"
                        SQL = SQL & ", INOUT"                           '입/외
                        SQL = SQL & ", REFVALUE"                        '참고치
                        SQL = SQL & ", CHARTNO"                         '챠트번호
                        SQL = SQL & ", PID                  " & vbCrLf  '병록번호(내원번호)"
                        SQL = SQL & ", PNAME"
                        SQL = SQL & ", PSEX"
                        SQL = SQL & ", PAGE"
                        SQL = SQL & ", PJUMIN"
                        SQL = SQL & ", SENDFLAG             " & vbCrLf  '전송구분(0:미전송,1:전송)"
                        SQL = SQL & ", SENDDATE"
                        SQL = SQL & ", EXAMUID"
                        SQL = SQL & ", EXAMTYPE"
                        SQL = SQL & ", EXAMNO"
                        SQL = SQL & ", HOSPITAL)            " & vbCrLf
                        SQL = SQL & " VALUES (              " & vbCrLf
                        SQL = SQL & 0
                        SQL = SQL & ",'" & Format(Now, "yyyymmddhhmmss") & "'"
                        SQL = SQL & ",'" & XMLInData.CommDate & "'"
                        SQL = SQL & ",'" & gHOSP.MACHCD & "'"
                        SQL = SQL & ",'" & strBarNum & "'                           " & vbCrLf
                        SQL = SQL & ",'" & XMLInData.ComExamID & "'"
                        SQL = SQL & ",''"
                        SQL = SQL & ",'" & XMLInData.ExamID & "'"
                        SQL = SQL & ",''"
                        SQL = SQL & ",'" & Trim(RS_L.Fields("TESTNAME").Value) & "' " & vbCrLf
                        SQL = SQL & ",'" & XMLInData.Specimen & "'"                             '검체유형
                        SQL = SQL & ",'" & XMLInData.IOFlag & "'"
                        SQL = SQL & ",'" & XMLInData.Reference & "'"
                        SQL = SQL & ",'" & XMLInData.ChartNo & "'"
                        SQL = SQL & ",'" & XMLInData.PatNo & "'                     " & vbCrLf
                        SQL = SQL & ",'" & XMLInData.PatName & "'"
                        SQL = SQL & ",'" & mPatient.SEX & "'"
                        SQL = SQL & ",'" & mPatient.AGE & "'"
                        SQL = SQL & ",'" & strJumin & "'"
                        SQL = SQL & ",'0'                                           " & vbCrLf  '전송구분(0:미전송,1:전송)
                        SQL = SQL & ",''"
                        SQL = SQL & ",'" & gHOSP.USERID & "'"
                        SQL = SQL & ",'" & gHOSP.PARTCD & "'"
                        SQL = SQL & ",'" & XMLInData.ExamNo & "'"
                        SQL = SQL & ",'" & XMLInData.HospCode & "')                 " & vbCrLf
                            
                    End If
                    
                    RS_C.Close
                    If Not DBExec(AdoCn_Local, SQL) Then
                        Call SetSQLData("저장에러", SQL, "A")
                    End If
                    
                End If
                RS_L.Close
            Next
            
        End If
    End If
    
    '5. 조회기간의 데이터를 불러온다.
    SQL = ""
    SQL = SQL & "Select DISTINCT "
    SQL = SQL & "       SAVESEQ                 " & vbCrLf
    SQL = SQL & "     , HOSPDATE                " & vbCrLf
    SQL = SQL & "     , INOUT                   " & vbCrLf
    SQL = SQL & "     , CHARTNO                 " & vbCrLf
    SQL = SQL & "     , BARCODE                 " & vbCrLf
    SQL = SQL & "     , PID                     " & vbCrLf
    SQL = SQL & "     , PNAME                   " & vbCrLf
    SQL = SQL & "     , PSEX                    " & vbCrLf
    SQL = SQL & "     , PAGE                    " & vbCrLf
    SQL = SQL & "     , PJUMIN                  " & vbCrLf
    SQL = SQL & "     , COUNT(EXAMCODE) AS CNT  " & vbCrLf
    SQL = SQL & "  From UB_PATRESULT                                            " & vbCrLf
    SQL = SQL & " Where HOSPDATE Between '" & pFrom & "' AND '" & pTo & "'      " & vbCrLf
    SQL = SQL & "   And EXAMCODE IN (" & gAllTestCd & ")                        " & vbCrLf
    If frmInterface.chkSave.Value = "0" Then
        SQL = SQL & "   And (RESULT = '' OR RESULT IS NULL) " & vbCrLf
        SQL = SQL & "   AND SENDFLAG = '0'                  " & vbCrLf
    End If
    SQL = SQL & " Group By SAVESEQ,HOSPDATE,INOUT,CHARTNO,BARCODE,PID,PNAME,PSEX,PAGE,PJUMIN " & vbCrLf
    SQL = SQL & " Order by HOSPDATE,PID,PNAME,BARCODE                               " & vbCrLf
            
    Call SetSQLData("워크조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn_Local.CursorLocation = adUseClient
    Set RS = AdoCn_Local.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        SPD.MaxRows = 0
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strBarcode = GetText(SPD, i, colBARCODE)
                    strPID = GetText(SPD, i, colPID)
                    If Trim(RS("HOSPDATE")) = strHospDate And Trim(RS("BARCODE")) = strBarcode And Trim(RS("PID")) = strPID Then
                        blnSame = True
                    End If
                Next
                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows
                        
                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    Select Case Trim(Trim(RS.Fields("INOUT")) & "")
                        Case "1":   SetText SPD, "외래", intRow, colINOUT
                        Case "0":   SetText SPD, "입원", intRow, colINOUT
                        Case Else:  SetText SPD, Trim(Trim(RS.Fields("INOUT")) & ""), intRow, colINOUT
                    End Select
                    
                    SetText SPD, IIf(Trim(RS.Fields("BARCODE")) & "" = "", Trim(RS.Fields("CHARTNO")), Trim(RS.Fields("BARCODE"))), intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("PJUMIN")) & "", intRow, colPJUMIN
                    SetText SPD, Trim(RS.Fields("PSEX")) & "", intRow, colPSEX
                    SetText SPD, Trim(RS.Fields("PAGE")) & "", intRow, colPAGE
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    SetText SPD, mOrder.OCNT, intRow, colOCNT
                End If
            End With
            
            blnSame = False
        
            DoEvents
            
            RS.MoveNext
        Loop
    Else
'        If gWORKPOS = "P" Then
'            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
'            frmWorkList.chkAll.Value = "0"
'        End If
    End If
    
    RS.Close
     
    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetWorkList_UBCARE" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    
    frmErrMsg.Show vbModal
    
    Screen.MousePointer = 0
    
End Sub

'MS-SQL
Public Sub GetWorkList_LABSPEAR(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As Object)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean

    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strTestCds  As String
    Dim strItems    As String
    
On Error GoTo ErrHandle

    Screen.MousePointer = 11
    blnSame = False
    strTestCds = ""

    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       CONVERT(NVARCHAR(50),M.접수일자,112)    AS HOSPDATE                                 " & vbCrLf
    SQL = SQL & "     , M.접수번호                              AS PID                                      " & vbCrLf
    SQL = SQL & "     , M.차트번호                              AS CHARTNO                                  " & vbCrLf
    SQL = SQL & "     , M.성명                                  AS PNAME                                    " & vbCrLf
    SQL = SQL & "     , M.성별                                  AS SEX                                      " & vbCrLf
    SQL = SQL & "     , M.나이                                  AS AGE                                      " & vbCrLf
    SQL = SQL & "     , M.거래처명                              AS DEPT                                     " & vbCrLf
    SQL = SQL & "     , E.검사코드                              AS ITEM                                     " & vbCrLf
    'SQL = SQL & "     , COUNT(E.검사코드)                       AS CNT                                      " & vbCrLf
    SQL = SQL & "  FROM VW_검사접수 M                                                                       " & vbCrLf
    SQL = SQL & "     , VW_검사결과 R                                                                       " & vbCrLf
    SQL = SQL & "     , VW_검사코드 E                                                                       " & vbCrLf
    SQL = SQL & " WHERE M.접수일자 BETWEEN CONVERT(DATE, '" & pFrom & "') AND CONVERT(DATE, '" & pTo & "')  " & vbCrLf
    SQL = SQL & "   AND M.접수일자      = R.접수일자                                                        " & vbCrLf
    SQL = SQL & "   AND M.접수번호      = R.접수번호                                                        " & vbCrLf
    SQL = SQL & "   AND R.검사코드      = E.검사코드                                                        " & vbCrLf
    SQL = SQL & "   AND E.학부코드      = '" & gHOSP.PARTCD & "'                                            " & vbCrLf    'U2
    SQL = SQL & "   AND E.검사코드      IN (" & gAllTestCd & ")                                             " & vbCrLf
    SQL = SQL & "   AND ISNULL(R.보고여부, 'N') <> 'Y'                                                      " & vbCrLf
    SQL = SQL & "   AND (R.결과값 IS NULL OR R.결과값 = '')                                                 " & vbCrLf
    'SQL = SQL & " GROUP BY M.접수일자,M.접수번호,M.차트번호,M.성명,M.성별,M.나이,M.거래처명                 " & vbCrLf
    
    Call SetSQLData("워크조회", SQL, "")

    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then

        SPD.MaxRows = 0

        Do Until RS.EOF
            With SPD
                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strBarcode = GetText(SPD, i, colBARCODE)
                    If Trim(RS("HOSPDATE")) = strHospDate And strBarcode = Trim(RS.Fields("HOSPDATE")) & PedLeftStr(Trim(RS.Fields("PID")), 5, "0") Then
                        blnSame = True
                        strItems = GetText(SPD, intRow, colITEMS)
                        strItems = strItems & GetTestNmS(Trim(RS.Fields("ITEM")))
                        SetText SPD, strItems, intRow, colITEMS
                    End If
                Next

                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows

                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & PedLeftStr(Trim(RS.Fields("PID")), 5, "0"), intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                    SetText SPD, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
                    'SetText SPD, Trim(RS.Fields("DEPT")) & "", intRow, colDEPT
                    'SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    'SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    
                    strItems = GetText(SPD, intRow, colITEMS)
                    strItems = strItems & GetTestNmS(Trim(RS.Fields("ITEM")))
                    SetText SPD, strItems, intRow, colITEMS
                    
                    '장비에서 오더요청이 안오는 배치오더용
                    Select Case gHOSP.MACHNM
                        Case "ACCESS2"
                            strItems = GetEquipExamCode_ACCESS2(gHOSP.MACHCD, "")
                            Call SetTag(SPD, strItems, intRow, colSTATE)
                        
                        Case "PPC300N"
                            strItems = GetEquipExamCode_PPC300N(gHOSP.MACHCD, "")
                            Call SetTag(SPD, strItems, intRow, colSTATE)
                    End Select
                    
                End If
                
            End With

            blnSame = False

            DoEvents

            RS.MoveNext
        Loop
    Else
        frmInterface.lblComStatus.Caption = "워크리스트 조회 대상자가 없습니다."
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetWorkList_AMIS" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    
    frmErrMsg.Show vbModal

End Sub

Public Sub GetWorkList_BIT70(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As Object)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean

    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strChartNo  As String
    Dim strTestCds  As String
    
On Error GoTo ErrHandle

    Screen.MousePointer = 11
    blnSame = False
    strTestCds = ""

    pTo = pTo & "235959"
    
'    SQL = SQL & "     , L.LABINSNUM     as      처방순서    " & vbCrLf
'    SQL = SQL & "     , L.LABSMPNAM     as      검체명      " & vbCrLf
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       L.LABODRDTE                     AS      HOSPDATE        " & vbCrLf
    SQL = SQL & "     , L.LABBARCOD                     AS      BARCODE         " & vbCrLf
    SQL = SQL & "     , L.LABCHTNUM                     AS      CHARTNO         " & vbCrLf
    SQL = SQL & "     , L.LABATTEND                     AS      PID             " & vbCrLf
    SQL = SQL & "     , M.MANADMFOR                     AS      INOUT           " & vbCrLf
    SQL = SQL & "     , M.MANRESNUM                     AS      JUMIN           " & vbCrLf
    SQL = SQL & "     , M.MANPATNAM                     AS      PNAME           " & vbCrLf
    SQL = SQL & "     , L.LABODRSTP                     AS      ORDCODE         " & vbCrLf
    SQL = SQL & "  FROM ME_LABDAT   L                                           " & vbCrLf
    SQL = SQL & "     , ME_DAT      D                                           " & vbCrLf
    SQL = SQL & "     , ME_MAN      M                                           " & vbCrLf
    SQL = SQL & " WHERE L.LABODRDTE BETWEEN '" & pFrom & "' AND '" & pTo & "'   " & vbCrLf
    SQL = SQL & "   AND L.LABKEYNUM     = D.DATKEYNUM                           " & vbCrLf      '-- 테이블연결키값
    SQL = SQL & "   AND L.LABATTEND     = D.DATATTEND                           " & vbCrLf      '-- 내원번호
    SQL = SQL & "   AND L.LABATTEND     = M.MANATTEND                           " & vbCrLf      '-- 내원번호
    SQL = SQL & "   AND L.LABCHTNUM     = D.DATCHTNUM                           " & vbCrLf      '-- 챠트번호
    SQL = SQL & "   AND L.LABCHTNUM     = M.MANCHTNUM                           " & vbCrLf      '-- 챠트번호
    SQL = SQL & "   AND L.LABODRDTE     = D.DATODRDTE                           " & vbCrLf      '-- 처방일자
    SQL = SQL & "   AND L.LABODRCOD     IN (" & gAllTestCd & ")                 " & vbCrLf      '-- 처방검사항목
    SQL = SQL & "   AND (L.LABCANCEL    = '' OR L.LABCANCEL IS NULL)            " & vbCrLf      '-- 취소여부
    SQL = SQL & "   AND (L.LABRESULT    = '' OR L.LABRESULT IS NULL)            " & vbCrLf      '-- 검사결과
    SQL = SQL & "   AND L.LABENDDEP     < '3'                                   " & vbCrLf      '-- 처리상태 (2:접수, 3:결과입력)
'    SQL = SQL & " ORDER BY L.LABODRDTE, L.LABCHTNUM, L.LABBARCOD, L.LABINSNUM   " & vbCrLf

    Call SetSQLData("워크조회", SQL, "")

    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then

        SPD.MaxRows = 0

        Do Until RS.EOF
            With SPD
                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strChartNo = GetText(SPD, i, colCHARTNO)
                    If Trim(RS("HOSPDATE")) = strHospDate And Trim(RS("CHARTNO")) = strChartNo Then
                        blnSame = True
                    End If
                Next

                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows

                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("JUMIN")) & "", intRow, colPAGE
                    Select Case Trim(Trim(RS.Fields("INOUT")) & "")
                        Case "A":   SetText SPD, "외래", intRow, colINOUT
                        Case "F":   SetText SPD, "입원", intRow, colINOUT
                        Case Else:  SetText SPD, "", intRow, colINOUT
                    End Select
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                End If
                
            End With

            blnSame = False

            DoEvents

            RS.MoveNext
        Loop
    Else
        frmInterface.lblComStatus.Caption = "워크리스트 조회 대상자가 없습니다."
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetWorkList_BIT70" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    
    frmErrMsg.Show vbModal

End Sub

Public Sub GetWorkList_BIT(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As Object)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean

    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strChartNo  As String
    Dim strTestCds  As String
    Dim strItems    As String
    
On Error GoTo ErrHandle

    Screen.MousePointer = 11
    blnSame = False
    strTestCds = ""

'    pFrom = pFrom & "235959"
'    pTo = pTo & "235959"
      
'    SQL = ""
'    SQL = SQL & "SELECT DISTINCT "
'    SQL = SQL & "        SUBSTRING(R.RESACPDTM,1,8) AS HOSPDATE                 " & vbCrLf
'    SQL = SQL & "     , R.RESSPMNUM                 AS BARCODE                  " & vbCrLf 'RESOCMNUM
'    SQL = SQL & "     , R.RESCHTNUM                 AS CHARTNO                  " & vbCrLf
'    SQL = SQL & "     , P.PBSPATNAM                 AS PNAME                    " & vbCrLf
'    SQL = SQL & "     , COUNT(R.RESLABCOD)          AS CNT                      " & vbCrLf
''    SQL = SQL & "      , P.PBSRESNUM                AS JUMIN                   " & vbCrLf
''    SQL = SQL & "      , P.PBSSEXTYP                AS SEX                     " & vbCrLf
'    SQL = SQL & "  FROM RESINF AS R                                             " & vbCrLf
'    SQL = SQL & "WHERE O.OCMACPDTM BETWEEN '" & pFrom & "' AND '" & pTo & "'    " & vbCrLf
'    SQL = SQL & "  AND R.RESCHTNUM = P.PBSCHTNUM                                " & vbCrLf
'    SQL = SQL & "  AND R.RESLABCOD IN (" & gAllTestCd & ")                      " & vbCrLf
'    SQL = SQL & "  AND (R.RESREPTYP IS NULL OR R.RESREPTYP <> 'F')              " & vbCrLf         '--  'I':중간 'F' 완료"
'    SQL = SQL & "  AND (R.RESRLTVAL = ''  OR R.RESRLTVAL IS NULL)               " & vbCrLf
'    SQL = SQL & " GROUP BY R.RESACPDTM,R.RESSPMNUM,R.RESCHTNUM,P.PBSPATNAM      " & vbCrLf



    SQL = ""
    SQL = SQL & " select " 'odrdtm     as HOSPDATE1,                                       "
    SQL = SQL & "      P.PbsPatNam AS PNAME                                     " & vbCrLf
    SQL = SQL & "    , P.PbsResNum AS PJUMIN                                    " & vbCrLf
    SQL = SQL & "    , W.OcmAcpDtm AS HOSPDATE                                  " & vbCrLf
    SQL = SQL & "    , R.ResOcmNum AS BARCODE                                   " & vbCrLf
    SQL = SQL & "    , W.OcmChtNum AS CHARTNO                                   " & vbCrLf
    'SQL = SQL & "    , W.OcmPatTyp AS 환자유형                                   " & vbcrlf
    'SQL = SQL & "    , W.OcmDepCod AS 진료과코드                                 " & vbcrlf
    'SQL = SQL & "    , D.DepKorNam AS 진료과명                                   " & vbcrlf
    'SQL = SQL & "    , W.OcmDtrCod AS 주치의                                     " & vbcrlf
    'SQL = SQL & "    , W.OcmWrdCod AS 병동코드                                   " & vbcrlf
    SQL = SQL & "    , COUNT(R.ResLabCod) AS COUNT                               " & vbCrLf
    'SQL = SQL & "    , (Select top 1 OdrEntDtm From OdrInf Where OdrOcmNum = W.OcmNum Order by odrseq desc) as 처방일시 " & vbcrlf
    SQL = SQL & "    From OcmInf AS W                                           " & vbCrLf
    SQL = SQL & "       , OdrInf AS O                                           " & vbCrLf
    SQL = SQL & "       , ResInf AS R                                           " & vbCrLf
    SQL = SQL & "       , PbsInf AS P                                           " & vbCrLf
    SQL = SQL & "       , depmst AS D                                           " & vbCrLf
    SQL = SQL & "       , LabMst AS E                                           " & vbCrLf
    SQL = SQL & "    where W.OcmNum    = R.ResOcmNum                            " & vbCrLf
    SQL = SQL & "      and W.OcmChtNum = P.PbsChtNum                            " & vbCrLf
    SQL = SQL & "      and W.OcmNum    = O.OdrOcmNum                            " & vbCrLf
    SQL = SQL & "      and W.OcmDepCod = D.DepCod                               " & vbCrLf
    SQL = SQL & "      and D.DepExpDte  >= (Select Convert(varchar(10),Getdate(),112)) " & vbCrLf
    SQL = SQL & "      And O.Odrdelflg = 'N'                                    " & vbCrLf
    SQL = SQL & "      and O.OdrSeq = R.ResOdrSeq                               " & vbCrLf
    SQL = SQL & "      and R.ResLabCod = E.LabCod                               " & vbCrLf
    SQL = SQL & "      and Upper(R.ResLabCod) IN (" & UCase(gAllTestCd) & ")           " & vbCrLf
    SQL = SQL & "      and W.OcmComStt Not In ('CN', 'CR','VC')                 " & vbCrLf
    SQL = SQL & "      and substring(O.odrdtm,1,8) Between '" & pFrom & "' and '" & pTo & "'   " & vbCrLf
    SQL = SQL & "      and (R.ResRltVal is null or R.ResRltVal = '')            " & vbCrLf
    SQL = SQL & " GROUP BY P.PbsPatNam, P.PbsResNum , W.OcmAcpDtm, R.ResOcmNum, W.OcmChtNum " & vbCrLf
      
      
    Call SetSQLData("워크조회", SQL, "")

    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then

        SPD.MaxRows = 0

        Do Until RS.EOF
            With SPD
                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strBarcode = GetText(SPD, i, colBARCODE)
                    If Trim(RS("HOSPDATE")) = strHospDate And Trim(RS("BARCODE")) = strBarcode Then
                        blnSame = True
                    End If
                Next

                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows

                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("COUNT")) & "", intRow, colOCNT
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    
                    'MsgBox strBarcode
                    'MsgBox gPatOrdCd
                    'MsgBox gHOSP.MACHNM
                    
                    If gPatOrdCd <> "" And gHOSP.MACHNM = "E411" Then
                        '-- 배치오더일 경우 필요함
                        strItems = GetEquipExamCode_E411(gHOSP.MACHCD, strBarcode)
                        
                     '   MsgBox strItems
                        
                        '-- 검사채널로 장비오더 만들기
                        If Trim(strItems) = "" Then
                            mOrder.NoOrder = True
                            mOrder.Order = ""
                            Call SetText(SPD, "오더없음", intRow, colSTATE)
                            Call SetText(SPD, "", intRow, colSPECIMEN)
                        Else
                            mOrder.NoOrder = False
                            mOrder.Order = strItems
                            Call SetText(SPD, "오더준비", intRow, colSTATE)
                            Call SetText(SPD, strItems, intRow, colSPECIMEN)
                            'MsgBox strItems
                        End If
                    End If
                    
                End If
                
            End With

            blnSame = False

            DoEvents

            RS.MoveNext
        Loop
    Else
        frmInterface.lblComStatus.Caption = "워크리스트 조회 대상자가 없습니다."
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetWorkList_BIT" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    
    frmErrMsg.Show vbModal

End Sub

Public Sub GetWorkList_MCC(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As Object)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean

    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strChartNo  As String
    Dim strTestCds  As String
    
On Error GoTo ErrHandle

    Screen.MousePointer = 11
    blnSame = False
    strTestCds = ""

    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       READING_YMD                 AS HOSPDATE                 " & vbCrLf
    SQL = SQL & "     , BCODE_NO                    AS BARCODE                  " & vbCrLf
'    SQL = SQL & "     , RECEPT_NO                  AS CHARTNO                  " & vbcrlf
    SQL = SQL & "     , PTNT_NO                     AS PID                      " & vbCrLf
    SQL = SQL & "     , PTNT_NM                     AS PNAME                    " & vbCrLf
    SQL = SQL & "     , AGE                         AS AGE                      " & vbCrLf
    SQL = SQL & "     , SEX                         AS SEX                      " & vbCrLf
    SQL = SQL & "     , IO_GB                       AS INOUT                    " & vbCrLf
    SQL = SQL & "     , COUNT(ORD_CD)               AS COUNT                    " & vbCrLf
    SQL = SQL & "  FROM LIS_INTERFACE1_V                                        " & vbCrLf
    SQL = SQL & " WHERE READING_YMD BETWEEN '" & pFrom & "' AND '" & pTo & "'   " & vbCrLf
    SQL = SQL & "   AND ORD_CD IN (" & gAllTestCd & ")                          " & vbCrLf
    SQL = SQL & "   AND STS_CD = '0'                                            " & vbCrLf    '0 접수, 1:결과전송
    SQL = SQL & " GROUP BY READING_YMD,BCODE_NO,PTNT_NO,PTNT_NM,AGE,SEX,IO_GB   " & vbCrLf
    SQL = SQL & " ORDER BY READING_YMD,PTNT_NO,BCODE_NO                         " & vbCrLf
    
    Call SetSQLData("워크조회", SQL, "")

    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then

        SPD.MaxRows = 0

        Do Until RS.EOF
            With SPD
                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strBarcode = GetText(SPD, i, colBARCODE)
                    If Trim(RS("HOSPDATE")) = strHospDate And Trim(RS("BARCODE")) = strBarcode Then
                        blnSame = True
                    End If
                Next

                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows

                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    'SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                    SetText SPD, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
                    SetText SPD, IIf(Trim(RS.Fields("INOUT")) & "" = "10", "입원", "외래"), intRow, colINOUT
                    SetText SPD, Trim(RS.Fields("COUNT")) & "", intRow, colRCNT
                    
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                End If
                
            End With

            blnSame = False

            DoEvents

            RS.MoveNext
        Loop
    Else
        frmInterface.lblComStatus.Caption = "워크리스트 조회 대상자가 없습니다."
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetWorkList_BIT" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    
    frmErrMsg.Show vbModal

End Sub


Public Sub GetWorkList_NEOSENSE(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As Object)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean

    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    
    Dim strHospYMD  As String
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strChartNo  As String
    Dim strTestCds  As String
    
    Dim strFY       As String
    Dim strFM       As String
    Dim strFD       As String
    Dim strTY       As String
    Dim strTM       As String
    Dim strTD       As String
    
    
On Error GoTo ErrHandle

    Screen.MousePointer = 11
    blnSame = False
    strTestCds = ""
    
    strFY = Mid(pFrom, 1, 4)
    strFM = Mid(pFrom, 5, 2)
    strFD = Mid(pFrom, 7, 2)
    strTY = Mid(pTo, 1, 4)
    strTM = Mid(pTo, 5, 2)
    strTD = Mid(pTo, 7, 2)
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       a.챠트번호          AS  CHARTNO " & vbCrLf
    SQL = SQL & "     , a.진료년            AS  HOSPYY  " & vbCrLf
    SQL = SQL & "     , a.진료월            AS  HOSPMM  " & vbCrLf
    SQL = SQL & "     , a.진료일            AS  HOSPDD  " & vbCrLf
'    SQL = SQL & "     , c.진료상태                      " & vbCrLf
    SQL = SQL & "     , b.수진자명          AS  PNAME   " & vbCrLf
'    SQL = SQL & "     , b.챠트번호          AS  CHARTNO " & vbCrLf
    SQL = SQL & "     , b.주민등록번호      AS  JUMIN   " & vbCrLf
    SQL = SQL & "     , COUNT(a.처방코드)     AS  COUNT   " & vbCrLf
    SQL = SQL & "  FROM TB_검사항목 a                   " & vbCrLf
    SQL = SQL & "     , TB_인적사항 b                   " & vbCrLf
    SQL = SQL & "     , tb_진료기본 c                   " & vbCrLf
    SQL = SQL & "     , tb_진료내역 d                   " & vbCrLf
    SQL = SQL & " WHERE a.진료년 Between '" & strFY & "' And '" & strTY & "'    " & vbCrLf
    SQL = SQL & "   And a.진료월 Between '" & strFM & "' And '" & strTM & "'    " & vbCrLf
    SQL = SQL & "   And a.진료일 Between '" & strFD & "' And '" & strTD & "'    " & vbCrLf
    SQL = SQL & "   And a.처방번호 > 0                                          " & vbCrLf
    'SQL = SQL & "   And c.진료상태 in( '1','5','6','7','8','9')                 " & vbCrLf  '6:DC처방
    'SQL = SQL & "   And c.진료상태 in( '1','5','7','8','9')                 " & vbCrLf  '6:DC처방
    SQL = SQL & "   And c.진료상태 in( '1','5','7','8')                 " & vbCrLf  '6:DC처방, 9:취소
    SQL = SQL & "   And a.처방코드 + '|' + a.서브코드 IN (" & gAllTestCd & ")         " & vbCrLf
    SQL = SQL & "   And (a.검사결과 is null or a.검사결과 = '')                 " & vbCrLf
    SQL = SQL & "   And d.수정구분 <> '1'                                       " & vbCrLf '//진료내역삭제
    SQL = SQL & "   And a.진료년    = c.진료년          " & vbCrLf
    SQL = SQL & "   And a.진료월    = c.진료월          " & vbCrLf
    SQL = SQL & "   And a.진료일    = c.진료일          " & vbCrLf
    SQL = SQL & "   And a.챠트번호  = c.챠트번호        " & vbCrLf
    SQL = SQL & "   And a.챠트번호  = b.챠트번호        " & vbCrLf
    SQL = SQL & "   And a.진료년    = d.진료년          " & vbCrLf
    SQL = SQL & "   And a.진료월    = d.진료월          " & vbCrLf
    SQL = SQL & "   And a.진료일    = d.진료일          " & vbCrLf
    SQL = SQL & "   And a.챠트번호  = d.챠트번호        " & vbCrLf
    SQL = SQL & "   And a.처방코드  = d.처방코드        " & vbCrLf
    '-- 2020-01-31 묶음처방은 서브코드가 없다
    'SQL = SQL & "   And a.서브코드  = d.서브코드        " & vbCrLf
    SQL = SQL & " GROUP BY a.챠트번호, a.진료년, a.진료월, a.진료일, b.수진자명, b.주민등록번호 " & vbCrLf
    SQL = SQL & " Order By a.진료년, a.진료월, a.진료일, b.수진자명 " & vbCrLf
      
    Call SetSQLData("워크조회", SQL, "")

    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then

        SPD.MaxRows = 0

        Do Until RS.EOF
            With SPD
                strHospYMD = Trim(RS.Fields("HOSPYY")) & Trim(RS.Fields("HOSPMM")) & Trim(RS.Fields("HOSPDD")) & ""
                
                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strBarcode = GetText(SPD, i, colBARCODE)
                    strChartNo = GetText(SPD, i, colCHARTNO)
                    
                    If strHospYMD = strHospDate And Trim(RS("CHARTNO")) = strChartNo Then
                        blnSame = True
                    End If
                Next

                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows

                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, strHospYMD, intRow, colHOSPDATE
                    'SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    
                    'Call CalAgeSex(Trim(RS.Fields("JUMIN")) & "", Format(Now, "yyyy/mm/dd"))
                   
                    'SetText SPD, mPatient.SEX, intRow, colPSEX
                    'SetText SPD, mPatient.AGE, intRow, colPAGE
                    SetText SPD, Trim(RS.Fields("COUNT")) & "", intRow, colRCNT
                    
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                End If
                
            End With

            blnSame = False

            DoEvents

            RS.MoveNext
        Loop
    Else
        frmInterface.lblComStatus.Caption = "워크리스트 조회 대상자가 없습니다."
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetWorkList_NEOSOFT" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    
    frmErrMsg.Show vbModal

End Sub

Public Sub GetWorkList_EASYS(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As Object, Optional ByVal pSave As Boolean = False)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim i           As Integer
    Dim intRow      As Integer
    Dim srJumin     As String
    Dim strAge      As String
    Dim strJuminSex As String
    
On Error GoTo ErrHandle

    Screen.MousePointer = 11
    blnSame = False
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       a.ACC_YMD       AS HOSPDATE     " & vbCrLf
    SQL = SQL & "     , a.RECEPT_NO     AS BARCODE      " & vbCrLf
    SQL = SQL & "     , a.PTNT_NO       AS PID          " & vbCrLf
    SQL = SQL & "     , c.PTNT_NM       AS PNAME        " & vbCrLf
    SQL = SQL & "     , c.BIRTH_YMD     AS AGE          " & vbCrLf
    SQL = SQL & "     , c.SEX           AS SEX          " & vbCrLf
    SQL = SQL & "     , COUNT(a.ORD_CD) AS CNT          " & vbCrLf
    SQL = SQL & "  FROM H3LAB_RESULT a                  " & vbCrLf
    SQL = SQL & "     , H1OPDIN b                       " & vbCrLf
    SQL = SQL & "     , HZ_MST_PTNT c                   " & vbCrLf
    SQL = SQL & " WHERE a.ACC_YMD between '" & pFrom & "' AND '" & pTo & "' " & vbCrLf
    SQL = SQL & "   AND a.ORD_CD IN (" & gAllTestCd & ")                    " & vbCrLf
    If pSave = False Then
        SQL = SQL & "   AND a.STS_CD     = 'A'                                  " & vbCrLf 'A:접수, R:결과전송
    End If
    SQL = SQL & "   AND (a.RESULT_NM = '' OR a.RESULT_NM IS NULL)            " & vbCrLf
'    SQL = SQL & "   AND a.SUTAK_CD   = ''                                   " & vbCrLf
    SQL = SQL & "   AND a.RECEPT_NO  = b.RECEPT_NO                          " & vbCrLf
    SQL = SQL & "   AND a.PTNT_NO    = c.PTNT_NO                            " & vbCrLf
    SQL = SQL & " GROUP BY a.ACC_YMD, a.RECEPT_NO, a.PTNT_NO, c.PTNT_NM,c.BIRTH_YMD,c.SEX " & vbCrLf
    SQL = SQL & " ORDER BY a.ACC_YMD, a.PTNT_NO " & vbCrLf
      
    Call SetSQLData("워크조회", SQL, "")

    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then

        SPD.MaxRows = 0

        Do Until RS.EOF
            With SPD
                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strBarcode = GetText(SPD, i, colBARCODE)
                    If Trim(RS("HOSPDATE")) = strHospDate And Trim(RS("BARCODE")) = strBarcode Then
                        blnSame = True
                    End If
                Next

                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows

                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                    SetText SPD, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
                    
                    Select Case Trim(RS.Fields("SEX")) & ""
                    Case "M"
                                If Mid(Trim(RS.Fields("AGE")) & "", 1, 1) = "1" Then
                                    strJuminSex = "1"
                                ElseIf Mid(Trim(RS.Fields("AGE")) & "", 1, 1) = "2" Then
                                    strJuminSex = "3"
                                End If
                    Case "F"
                                If Mid(Trim(RS.Fields("AGE")) & "", 1, 1) = "1" Then
                                    strJuminSex = "2"
                                ElseIf Mid(Trim(RS.Fields("AGE")) & "", 1, 1) = "2" Then
                                    strJuminSex = "4"
                                End If
                    End Select
                    
                    srJumin = Mid(Trim(RS.Fields("AGE")) & "", 3, 6) & strJuminSex & "000000"
                    
                    
                    Call CalAgeSex(srJumin, Format(Now, "yyyy/mm/dd"))
                    
                    SetText SPD, mPatient.SEX, intRow, colPSEX
                    SetText SPD, mPatient.AGE, intRow, colPAGE
                    
                    
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colRCNT
                    
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                End If
                
            End With

            blnSame = False

            DoEvents

            RS.MoveNext
        Loop
    Else
        '조회대상자 없음
        frmInterface.lblComStatus.Caption = "워크리스트 조회 대상자가 없습니다."
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetWorkList_EASYS" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    
    frmErrMsg.Show vbModal

End Sub


Public Sub GetWorkList_HANARO(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As Object)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean

    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strTestCds  As String
    
On Error GoTo ErrHandle

    Screen.MousePointer = 11
    blnSame = False
    strTestCds = ""

    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       ACPTNO  AS BARCODE              " & vbCrLf
    SQL = SQL & "     , CASU    AS CHARTNO              " & vbCrLf
    SQL = SQL & "     , WORKNO  AS PID                  " & vbCrLf
    SQL = SQL & "     , NAME    AS PNAME                " & vbCrLf
    SQL = SQL & "     , SEX     AS SEX                  " & vbCrLf
    SQL = SQL & "  From SZDAT01T                        " & vbCrLf
    SQL = SQL & " WHERE COMP    = '" & gHOSP.HOSPCD & "'" & vbCrLf
    SQL = SQL & "   AND WORKCD  = '" & gHOSP.MACHCD & "'" & vbCrLf
    SQL = SQL & "   AND ACPTNO  = '" & strBarcode & "' " & vbCrLf
    
    Call SetSQLData("워크조회", SQL, "")

    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then

        SPD.MaxRows = 0

        Do Until RS.EOF
            With SPD
                For i = 1 To SPD.DataRowCnt
                    strHospDate = GetText(SPD, i, colHOSPDATE)
                    strBarcode = GetText(SPD, i, colBARCODE)
                    If Trim(RS("HOSPDATE")) = strHospDate And Trim(RS("BARCODE")) = strBarcode Then
                        blnSame = True
                    End If
                Next

                If blnSame = False Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows

                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("SEX")) & "", intRow, colPSEX
                    SetText SPD, Trim(RS.Fields("AGE")) & "", intRow, colPAGE
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    
'                    If gWORKPOS = "P" Then
                        SetText SPD, frmWorkList.txtSeqNo.Text, intRow, colSEQNO
                        frmWorkList.txtSeqNo.Text = frmWorkList.txtSeqNo.Text + 1
'                    Else
'                        SetText SPD, frmInterface.txtSeqNo.Text, intRow, colSEQNO
'                        frmInterface.txtSeqNo.Text = frmInterface.txtSeqNo.Text + 1
'                    End If
                End If
                
            End With

            blnSame = False

            DoEvents

            RS.MoveNext
        Loop
        If gWORKPOS = "P" Then
'            frmWorkList.chkAll.Value = "1"
        End If
    Else
        If gWORKPOS = "P" Then
'            frmWorkList.lblStatus.Caption = ">> 조회 대상자가 없습니다."
'            frmWorkList.chkAll.Value = "0"
        End If
    End If

    RS.Close

    SPD.RowHeight(-1) = 12
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "Form_Load" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show vbModal

End Sub


'-- 결과저장용 키 가져오기
Function GetSampleSubITEM(ByVal pBarcode As String, ByVal pTestCd As String, Optional ByVal pRegDate As String, Optional ByVal pChartNo As String) As String

    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strRegDate      As String
    Dim strChartNo      As String
    Dim strInOut        As String
    
    Dim lngExamNo       As Long
    Dim strItems        As String
    Dim strSpcYY        As String
    Dim strSpcNo        As String
    
    GetSampleSubITEM = ""
        
    Select Case gEMR
        Case "JWINFO"
                
                SQL = ""
                SQL = SQL & "SELECT DISTINCT ORDERCODE AS ORDCODE       " & vbCrLf
                SQL = SQL & "  FROM SLA_LabMaster                       " & vbCrLf
                SQL = SQL & " WHERE SPECIMENNUM = '" & strBarcode & "'  " & vbCrLf
                SQL = SQL & "   AND LABCODE     = '" & pTestCd & "'     " & vbCrLf
                
                Set RS = AdoCn.Execute(SQL, , 1)
                If Not RS.EOF = True And Not RS.BOF = True Then
                    Do Until RS.EOF
                        GetSampleSubITEM = Trim(RS.Fields("ORDCODE")) & ""
                        RS.MoveNext
                    Loop
                End If
                
                RS.Close
    
        Case "KCHART"
                SQL = ""
                SQL = SQL & "SELECT DISTINCT                                                    " & vbCrLf
                SQL = SQL & "       L.진료검사ID                AS ORDCODE                      " & vbCrLf
                SQL = SQL & "     , L.진료지원ID                AS TESTSUBCODE                  " & vbCrLf
                SQL = SQL & "  FROM             TB_진료검사 L                                   " & vbCrLf
                SQL = SQL & "       INNER JOIN  TB_진료지원 J ON  (L.진료지원ID = J.진료지원ID) " & vbCrLf
                SQL = SQL & "       INNER JOIN  TB_진료일반 A ON  (J.진료일자   = A.진료일자    " & vbCrLf
                SQL = SQL & "                                AND   J.챠트번호   = A.챠트번호    " & vbCrLf
                SQL = SQL & "                                AND   J.진료번호   = A.진료번호)   " & vbCrLf
                SQL = SQL & " Where L.검체번호= '" & pBarcode & "'                              " & vbCrLf
                SQL = SQL & "   AND L.검사상태 < 5                                              " & vbCrLf
                SQL = SQL & "   AND (L.처방코드 + L.서브코드) = '" & pTestCd & "'               " & vbCrLf
    
                Call SetSQLData("SUB코드조회", SQL)
                
                Set RS = AdoCn.Execute(SQL, , 1)
                If Not RS.EOF = True And Not RS.BOF = True Then
                    Do Until RS.EOF
                        GetSampleSubITEM = Trim(RS.Fields("ORDCODE")) & "|" & Trim(RS.Fields("TESTSUBCODE"))
                        RS.MoveNext
                    Loop
                End If
                
                RS.Close
        
        Case "AMIS"
                
                SQL = ""
                SQL = SQL & "SELECT DISTINCT R.ORDERCODE           as ORDCODE   " & vbCrLf
                SQL = SQL & "  FROM REGISTINFOS O, RESULTOFNUM R, PATMST P      " & vbCrLf
                SQL = SQL & " WHERE O.ACPTDATE  = R.ACPTDATE                    " & vbCrLf
                SQL = SQL & "   AND O.PATID     = R.PATID                       " & vbCrLf
                SQL = SQL & "   AND O.ACPTSEQ   = R.ACPTSEQ                     " & vbCrLf
                SQL = SQL & "   AND O.PATID     = P.PATID                       " & vbCrLf
                SQL = SQL & "   AND R.SPCMNO    = '" & pBarcode & "'            " & vbCrLf
                SQL = SQL & "   AND R.RESULTITEMCODE = '" & pTestCd & "'        " & vbCrLf
                SQL = SQL & "   AND O.CLAS          = 4                         " & vbCrLf '임상병리
                SQL = SQL & "   AND R.RESULTFLAG    = 0                         " & vbCrLf
                SQL = SQL & " ORDER BY R.SPCMNO                                 " & vbCrLf
                
                Set RS = AdoCn.Execute(SQL, , 1)
                If Not RS.EOF = True And Not RS.BOF = True Then
                    Do Until RS.EOF
                        GetSampleSubITEM = Trim(RS.Fields("ORDCODE")) & ""
                        RS.MoveNext
                    Loop
                End If
                
                RS.Close
        
        Case "BIT70"
                
                SQL = ""
                SQL = SQL & "SELECT DISTINCT L.LABODRSTP as ORDCODE             " & vbCrLf
                SQL = SQL & "  FROM ME_LABDAT L, ME_DAT D                       " & vbCrLf
                SQL = SQL & " WHERE L.LABODRDTE  = '" & pRegDate & "'           " & vbCrLf
                SQL = SQL & "   AND L.LABCHTNUM  = '" & pChartNo & "'           " & vbCrLf
                SQL = SQL & "   AND L.LABKEYNUM  = D.DATKEYNUM                  " & vbCrLf    '-- 테이블연결키값
                SQL = SQL & "   AND L.LABATTEND  = D.DATATTEND                  " & vbCrLf    '-- 내원번호
                SQL = SQL & "   AND L.LABATTEND  = M.MANATTEND                  " & vbCrLf    '-- 내원번호 ???
                SQL = SQL & "   AND L.LABCHTNUM  = D.DATCHTNUM                  " & vbCrLf    '-- 챠트번호
                SQL = SQL & "   AND L.LABCHTNUM  = M.MANCHTNUM                  " & vbCrLf    '-- 챠트번호
                SQL = SQL & "   AND L.LABODRDTE  = D.DATODRDTE                  " & vbCrLf    '-- 처방일자
                SQL = SQL & "   AND L.LABODRCOD IN (" & gAllTestCd & ")         " & vbCrLf
                SQL = SQL & "   AND (L.LABCANCEL = '' OR L.LABCANCEL IS NULL)   " & vbCrLf    '-- 취소여부
                SQL = SQL & "   AND (L.LABRESULT = ''  OR L.LABRESULT IS NULL)  " & vbCrLf
                SQL = SQL & "   AND L.LABENDDEP < '3'                           " & vbCrLf    '-- 처리상태 (2:접수, 3:결과입력)
                    
                Set RS = AdoCn.Execute(SQL, , 1)
                If Not RS.EOF = True And Not RS.BOF = True Then
                    Do Until RS.EOF
                        GetSampleSubITEM = Trim(RS.Fields("ORDCODE")) & ""
                        RS.MoveNext
                    Loop
                End If
                
                RS.Close
        
        
        Case "EONM"
                SQL = ""
                SQL = SQL & "SELECT DISTINCT O.H141_SEQNO   AS SUBITEM      " & vbCrLf
                SQL = SQL & "  FROM TB_H141_LISTAKEBODY O,TB_A110_PATINFO P " & vbCrLf
                SQL = SQL & " Where P.A110_ChartNo    = O.H141_CHARTNO      " & vbCrLf
                SQL = SQL & "   AND O.H141_TSAMPLENO  = '" & pBarcode & "'  " & vbCrLf
                SQL = SQL & "   And O.H141_SUGACD     = '" & pTestCd & "'   " & vbCrLf
            
                Set RS = AdoCn.Execute(SQL, , 1)
                If Not RS.EOF = True And Not RS.BOF = True Then
                    Do Until RS.EOF
                        GetSampleSubITEM = Trim(RS.Fields("SUBITEM")) & ""
                        RS.MoveNext
                    Loop
                End If
                
                RS.Close
    End Select

End Function

'-- 검사자 ITEM 가져오기
Function GetSampleITEM(ByVal asRow As Long, ByVal SPD As Object) As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strRegDate      As String
    Dim strChartNo      As String
    Dim strInOut        As String
    
    Dim lngExamNo       As Long
    Dim strItems        As String
    Dim strSpcYY        As String
    Dim strSpcNo        As String
    
    Dim strYY           As String
    Dim strMM           As String
    Dim strDD           As String
    
    GetSampleITEM = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    strInOut = Trim(GetText(SPD, asRow, colINOUT))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    strYY = Mid(strRegDate, 1, 4)
    strMM = Mid(strRegDate, 5, 2)
    strDD = Mid(strRegDate, 7, 2)
        
    Select Case gEMR
        Case "BIT"
            SQL = ""
            SQL = SQL & " SELECT DISTINCT R.ResLabCod AS ITEM                   " & vbCrLf
            SQL = SQL & "   FROM RESINF AS R                                    " & vbCrLf
            SQL = SQL & " WHERE LTRIM(RTRIM(R.RESOCMNUM)) = '" & strBarcode & "'" & vbCrLf
            'SQL = SQL & " WHERE LTRIM(RTRIM(R.RESSPMNUM)) = '" & strBarcode & "'" & vbCrLf
            SQL = SQL & "   AND R.RESLABCOD IN (" & gAllTestCd & ")             " & vbCrLf
            SQL = SQL & "   AND (R.RESREPTYP IS NULL OR R.RESREPTYP <> 'F')     " & vbCrLf         '--  'I':중간 'F' 완료"
            SQL = SQL & "   AND (R.RESRLTVAL = ''  OR R.RESRLTVAL IS NULL)      " & vbCrLf
            'SQL = SQL & " Order By R.ResLabCod                                  " & vbCrLf
         
         Case "EASYS"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT ORD_CD AS ITEM                     " & vbCrLf
            SQL = SQL & "  FROM H3LAB_RESULT a, H1OPDIN b, HZ_MST_PTNT c    " & vbCrLf
            SQL = SQL & " WHERE a.ACC_YMD   = '" & strRegDate & "'          " & vbCrLf
            SQL = SQL & "   AND a.RECEPT_NO = '" & strBarcode & "'          " & vbCrLf
            SQL = SQL & "   AND a.ORD_CD IN (" & gAllTestCd & ")            " & vbCrLf
'            SQL = SQL & "   AND a.STS_CD    = 'A'                           " & vbCrLf 'A:접수, R:결과전송
'            SQL = SQL & "   AND a.SUTAK_CD  = ''                            " & vbCrLf
            SQL = SQL & "   AND (a.RESULT_NM = '' OR a.RESULT_NM IS NULL)            " & vbCrLf
            SQL = SQL & "   AND a.RECEPT_NO = b.RECEPT_NO                   " & vbCrLf
            SQL = SQL & " ORDER BY ORD_CD                                   " & vbCrLf
        
        Case "NEOSENSE"
            SQL = ""
            SQL = SQL & "SELECT (처방코드 + 서브코드) as ITEM               " & vbCrLf
            SQL = SQL & "  FROM TB_검사항목                                 " & vbCrLf
            SQL = SQL & " WHERE 진료년 = '" & strYY & "'                    " & vbCrLf
            SQL = SQL & "   And 진료월 = '" & strMM & "'                    " & vbCrLf
            SQL = SQL & "   And 진료일 = '" & strDD & "'                    " & vbCrLf
            SQL = SQL & "   And 처방번호 > 0                                " & vbCrLf
            SQL = SQL & "   And 처방코드 + '|' + 서브코드 IN (" & gAllTestCd & ") " & vbCrLf
            SQL = SQL & "   And (검사결과 is null or 검사결과 = '')         " & vbCrLf

        Case "AMIS"
            SQL = ""
            SQL = SQL & "SELECT R.RESULTITEMCODE as ITEM                    " & vbCr
            SQL = SQL & "  FROM registinfos O, resultofnum R                " & vbCr
            SQL = SQL & " WHERE O.acptdate = R.acptdate                     " & vbCr
            SQL = SQL & "   AND R.SPCMNO = '" & strBarcode & "'             " & vbCr
            SQL = SQL & "   AND O.patid = R.patid                           " & vbCr
            SQL = SQL & "   AND O.acptseq = R.acptseq                       " & vbCr
            SQL = SQL & "   AND O.CLAS = 4                                  " & vbCr '임상병리
'            SQL = SQL & "   AND R.RESULTFLAG = 0                            " & vbCr
            SQL = SQL & "   AND (R.NUMRESULTVAL = '' OR R.NUMRESULTVAL IS NULL)     " & vbCrLf
'            SQL = SQL & "   AND R.ORDERCODE IN (" & gAllOrdCd & ")          " & vbCr
            SQL = SQL & "   AND R.RESULTITEMCODE in (" & gAllTestCd & ")    " & vbCr
            SQL = SQL & "  ORDER BY R.RESULTITEMCODE                        " & vbCr
        
        Case "BIGUBCARE"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT i.IntLabCod + cast(IntLabseq as varchar(3)) AS ITEM "
            SQL = SQL & "  from interfacedb..IntRst i, aphdb..rstinf r " & vbCr
            SQL = SQL & " WHERE r.RstOdrStt not in ('OC') " & vbCr
            SQL = SQL & "   AND (r.rstrstval = '' or rstrstval is null)" & vbCr
            'If gHOSP.MACHNM <> "HITACHI7080" Then
                SQL = SQL & "   AND i.intodrtyp = '" & gHOSP.PARTCD & "'" & vbCr  ''HEMO'
            'End If
            SQL = SQL & "   AND i.IntOdrDte = '" & strRegDate & "'" & vbCr
            SQL = SQL & "   AND i.IntLabNum = '" & strBarcode & "'" & vbCr
            SQL = SQL & "   AND i.IntChtNum = '" & strChartNo & "'" & vbCr
'            SQL = SQL & "   AND i.IntLabCod IN (" & gAllTestCd & ")" & vbCr
            SQL = SQL & "   AND i.IntLabCod + cast(IntLabseq as varchar(3)) IN (" & gAllTestCd & ")" & vbCr
            SQL = SQL & "   AND i.intlabnum = r.rstlabnum" & vbCr
            SQL = SQL & "   AND i.intodrdte = r.rstodrdte" & vbCr
            SQL = SQL & "   AND i.intlabseq = r.rstlabseq" & vbCr
            SQL = SQL & "   AND i.intlabcod = r.rstodrcod" & vbCr
        
        Case "BIT70"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT L.LABODRCOD as ITEM                " & vbCrLf
            SQL = SQL & "  FROM ME_LABDAT L, ME_DAT D                       " & vbCrLf
            SQL = SQL & " WHERE L.LABODRDTE  = '" & strRegDate & "'         " & vbCrLf
            SQL = SQL & "   AND L.LABCHTNUM  = '" & strChartNo & "'         " & vbCrLf
            SQL = SQL & "   AND L.LABKEYNUM  = D.DATKEYNUM                  " & vbCrLf    '-- 테이블연결키값
            SQL = SQL & "   AND L.LABATTEND  = D.DATATTEND                  " & vbCrLf    '-- 내원번호
            SQL = SQL & "   AND L.LABATTEND  = M.MANATTEND                  " & vbCrLf    '-- 내원번호 ???
            SQL = SQL & "   AND L.LABCHTNUM  = D.DATCHTNUM                  " & vbCrLf    '-- 챠트번호
            SQL = SQL & "   AND L.LABCHTNUM  = M.MANCHTNUM                  " & vbCrLf    '-- 챠트번호
            SQL = SQL & "   AND L.LABODRDTE  = D.DATODRDTE                  " & vbCrLf    '-- 처방일자
            SQL = SQL & "   AND L.LABODRCOD IN (" & gAllTestCd & ")         " & vbCrLf
            SQL = SQL & "   AND (L.LABCANCEL = '' OR L.LABCANCEL IS NULL)   " & vbCrLf    '-- 취소여부
            SQL = SQL & "   AND (L.LABRESULT = ''  OR L.LABRESULT IS NULL)  " & vbCrLf
            SQL = SQL & "   AND L.LABENDDEP < '3'                           " & vbCrLf    '-- 처리상태 (2:접수, 3:결과입력)
'            SQL = SQL & " Order By L.LABODRCOD                              " & vbCrLf
        
        Case "EONM"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT O.H141_SUGACD AS ITEM              " & vbCrLf
            SQL = SQL & "  FROM TB_H141_LISTAKEBODY O, TB_A110_PATINFO P    " & vbCrLf
            SQL = SQL & " Where P.A110_ChartNo = O.H141_CHARTNO             " & vbCrLf
            SQL = SQL & "   AND O.H141_TSAMPLENO  = '" & strBarcode & "'    " & vbCrLf
            'SQL = SQL & "   AND O.H141_NOTYYN = 'N'                         " & vbCrlf
            SQL = SQL & "   AND O.H141_NOTYYN       IN ('N','T')                 " & vbCrLf '결과대기:T
            SQL = SQL & "   And O.H141_SUGACD in (" & gAllTestCd & ")       " & vbCrLf
            SQL = SQL & " ORDER BY O.H141_SUGACD                            " & vbCrLf
        
        
        Case "GINUS"
            SQL = ""
            SQL = SQL & "SELECT /*+ INDEX(rslt scrrslth_ux1) INDEX (coif scccoifm_ix1) */" & vbCrLf
            SQL = SQL & "       rslt.cd as ITEM                                         " & vbCrLf
            SQL = SQL & "  FROM scrrslth rslt                                           " & vbCrLf
            SQL = SQL & "     , scccoifm coif                                           " & vbCrLf
            SQL = SQL & "     , scccodem codm                                           " & vbCrLf
            SQL = SQL & "     , scrprexh prex                                           " & vbCrLf
            SQL = SQL & "     , mosxpslh xpsl                                           " & vbCrLf
            SQL = SQL & "     , pmcptbsm ptbs                                           " & vbCrLf
            SQL = SQL & "WHERE rslt.hos_org_no   = '" & gHOSP.HOSPCD & "'               " & vbCrLf
            SQL = SQL & "  AND rslt.smp_no       = '" & strBarcode & "'                 " & vbCrLf
            SQL = SQL & "  AND rslt.exam_stus  IN ('0','1','2')                         " & vbCrLf
            SQL = SQL & "  AND coif.hos_org_no   = rslt.hos_org_no                      " & vbCrLf
            'SQL = SQL & "  AND coif.exam_cd      = rslt.cd                              " & vbCrlf
            SQL = SQL & "  AND SUBSTR(prex.acp_dt,1,8) BETWEEN coif.fr_dt AND coif.to_dt" & vbCrLf
            SQL = SQL & "  AND SUBSTR(prex.acp_dt,1,8) BETWEEN codm.fr_dt AND codm.to_dt" & vbCrLf
            SQL = SQL & "  AND coif.exam_mach_cd = '" & gHOSP.MACHCD & "'               " & vbCrLf
            SQL = SQL & "  AND codm.hos_org_no   = coif.hos_org_no                      " & vbCrLf
            SQL = SQL & "  AND codm.typ_cd       = '02'                                 " & vbCrLf
            SQL = SQL & "  AND codm.cd           = coif.spc_cd                          " & vbCrLf
            SQL = SQL & "  AND prex.hos_org_no   = rslt.hos_org_no                      " & vbCrLf
            SQL = SQL & "  AND prex.smp_no       = rslt.smp_no                          " & vbCrLf
            SQL = SQL & "  AND prex.prcp_seq     = rslt.prcp_seq                        " & vbCrLf
            SQL = SQL & "  AND prex.exam_seq     = rslt.exam_seq                        " & vbCrLf
            SQL = SQL & "  AND xpsl.hos_org_no   = prex.hos_org_no                      " & vbCrLf
            SQL = SQL & "  AND xpsl.smp_no       = prex.smp_no                          " & vbCrLf
            SQL = SQL & "  AND xpsl.acp_no       = prex.prcp_seq                        " & vbCrLf
            SQL = SQL & "  AND xpsl.prcp_typ_cd IN ('O','C')                            " & vbCrLf
            SQL = SQL & "  AND ptbs.hos_org_no   = prex.hos_org_no                      " & vbCrLf
            SQL = SQL & "  AND ptbs.pt_no        = prex.pt_no                           " & vbCrLf
        
        Case "HWASAN"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT T.TESTCD as ITEM           " & vbCrLf
            SQL = SQL & "  FROM TC201 O, TC301 T                    " & vbCrLf
            SQL = SQL & " WHERE O.SPCNO = T.SPCNO                   " & vbCrLf
            SQL = SQL & "   AND O.SPCNO = '" & strBarcode & "'      " & vbCrLf
            SQL = SQL & "   And T.TESTCD in (" & gAllTestCd & ")    " & vbCrLf
            SQL = SQL & " Order By T.TESTCD                         " & vbCrLf
        
        Case "JAINCOM"
            SQL = ""
            SQL = SQL & "SELECT DiSTINCT b.SCP42SUGACD as ITEM                  " & vbCrLf
            SQL = SQL & "  FROM JAIN_SCP.SCPRST41 a, JAIN_SCP.SCPRST42 b        " & vbCrLf
            SQL = SQL & " WHERE a.SCP41PCODE    = b.SCP42PCODE                  " & vbCrLf
            SQL = SQL & "   AND a.SCP41JDATE    = b.SCP42JDATE                  " & vbCrLf
            SQL = SQL & "   AND a.SCP41SID      = b.SCP42SID                    " & vbCrLf
            SQL = SQL & "   AND a.SCP41SPMNO2   = b.SCP42SPMNO2                 " & vbCrLf
            SQL = SQL & "   AND a.SCP41SPMNO2   = '" & strBarcode & "'          " & vbCrLf
            SQL = SQL & "   AND b.SCP42SUGACD  IN (" & gAllTestCd & ")          " & vbCrLf
            SQL = SQL & "   AND (b.SCP42RESULT IS NULL OR b.SCP42RESULT = '')   " & vbCrLf
            SQL = SQL & " ORDER BY b.SCP42SUGACD                                " & vbCrLf
        
        Case "JWINFO"
            'AND ORDERCODE IN (" & gAllOrdCd & ") " & vbCrlf
            SQL = ""
            SQL = SQL & "SELECT DISTINCT LABCODE AS ITEM            " & vbCrLf
            SQL = SQL & "   FROM SLA_Labresult                      " & vbCrLf
            SQL = SQL & " WHERE LABCODE IN (" & gAllTestCd & ")     " & vbCrLf
            SQL = SQL & "   AND RECEIPTDATE = '" & strRegDate & "'  " & vbCrLf
            SQL = SQL & "   AND SPECIMENNUM = '" & strBarcode & "'  " & vbCrLf
            'SQL = SQL & "   AND JSTATUS < '3'                      " & vbCrlf
            SQL = SQL & " ORDER BY LABCODE                          " & vbCrLf
        
        Case "KOMAIN"
            SQL = ""
        
        Case "KCHART"
'            SQL = SQL & "    AND L.검사종류 = '" & gHOSP.LABCD & "'" & vbCrlf
            SQL = ""
            SQL = SQL & "SELECT DISTINCT (L.처방코드 + L.서브코드) AS ITEM                  " & vbCrLf
            SQL = SQL & "  FROM             TB_진료검사 L                                   " & vbCrLf
            SQL = SQL & "       INNER JOIN  TB_진료지원 J ON  (L.진료지원ID = J.진료지원ID) " & vbCrLf
            SQL = SQL & "       INNER JOIN  TB_진료일반 A ON  (J.진료일자   = A.진료일자    " & vbCrLf
            SQL = SQL & "                                AND   J.챠트번호   = A.챠트번호    " & vbCrLf
            SQL = SQL & "                                AND   J.진료번호   = A.진료번호)   " & vbCrLf
            SQL = SQL & " Where L.검체번호= '" & strBarcode & "'                            " & vbCrLf
            SQL = SQL & "   AND L.검사상태 < 5                                              " & vbCrLf
            SQL = SQL & "   AND L.처방코드 + L.서브코드 IN (" & gAllTestCd & ")             " & vbCrLf
'            SQL = SQL & " ORDER BY L.처방코드, L.서브코드                                   " & vbCrLf
            
        Case "KYU"
            SQL = ""
            
        
        Case "SANSOFT"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT "
            SQL = SQL & "       E.검사코드                              AS ITEM     " & vbCrLf
            SQL = SQL & "  FROM VW_검사접수 M, VW_검사결과 R, VW_검사코드 E         " & vbCrLf
            SQL = SQL & " WHERE M.접수일자 = CONVERT(DATE, '" & strRegDate & "')    " & vbCrLf
            SQL = SQL & "   AND M.접수일자 = R.접수일자                             " & vbCrLf
            SQL = SQL & "   AND M.접수번호 = R.접수번호                             " & vbCrLf
            SQL = SQL & "   AND R.검사코드 = E.검사코드                             " & vbCrLf
            SQL = SQL & "   AND m.접수번호 = '" & strPatID & "'                     " & vbCrLf
            SQL = SQL & "   AND E.학부코드 = '" & gHOSP.PARTCD & "'                 " & vbCrLf
            SQL = SQL & "   AND E.검사코드 IN (" & gAllTestCd & ")                  " & vbCrLf
            SQL = SQL & "   AND IsNull(R.보고여부, 'N') <> 'Y'                      " & vbCrLf
            SQL = SQL & "   AND (R.결과값 is null or R.결과값 = '')                 " & vbCrLf
        
        Case "LABSPEAR" 'PHILL
            SQL = ""
            SQL = SQL & "SELECT DISTINCT "
            SQL = SQL & "       E.검사코드                              AS ITEM     " & vbCrLf
            SQL = SQL & "  FROM VW_검사접수 M, VW_검사결과 R, VW_검사코드 E         " & vbCrLf
            SQL = SQL & " WHERE M.접수일자 = CONVERT(DATE, '" & strRegDate & "')    " & vbCrLf
            SQL = SQL & "   AND M.접수일자 = R.접수일자                             " & vbCrLf
            SQL = SQL & "   AND M.접수번호 = R.접수번호                             " & vbCrLf
            SQL = SQL & "   AND R.검사코드 = E.검사코드                             " & vbCrLf
            SQL = SQL & "   AND m.접수번호 = '" & strPatID & "'                     " & vbCrLf
            SQL = SQL & "   AND E.학부코드 = '" & gHOSP.PARTCD & "'                 " & vbCrLf
            SQL = SQL & "   AND E.검사코드 IN (" & gAllTestCd & ")                  " & vbCrLf
            SQL = SQL & "   AND IsNull(R.보고여부, 'N') <> 'Y'                      " & vbCrLf
            SQL = SQL & "   AND (R.결과값 is null or R.결과값 = '')                 " & vbCrLf
            
        Case "MCC"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT ORD_CD AS ITEM             " & vbCrLf
            SQL = SQL & "  FROM LIS_INTERFACE1_V                    " & vbCrLf
            SQL = SQL & " WHERE READING_YMD = '" & strRegDate & "'  " & vbCrLf
            SQL = SQL & "   AND BCODE_NO    = '" & strBarcode & "'  " & vbCrLf
            SQL = SQL & "   AND ORD_CD IN (" & gAllTestCd & ")      " & vbCrLf
            SQL = SQL & " ORDER BY ORD_CD                           " & vbCrLf
        
        Case "MEDICHART"
            SQL = ""
            SQL = SQL & "Select DISTINCT (a.처방코드 + a.서브코드)      AS ITEM     " & vbCrLf
            SQL = SQL & "  From TB_검사항목 a, TB_진료기본 c                        " & vbCrLf
            SQL = SQL & " Where a.챠트번호 = '" & strChartNo & "'                   " & vbCrLf
            SQL = SQL & "   And a.처방번호 > 0                                      " & vbCrLf
            SQL = SQL & "   And c.진료상태 IN ('1','5','6','7','8','9')             " & vbCrLf
            SQL = SQL & "   And (a.처방코드 + a.서브코드) IN (" & gAllTestCd & ")   " & vbCrLf
            SQL = SQL & "   And (a.검사결과 IS NULL OR a.검사결과 = '')             " & vbCrLf
            SQL = SQL & "   And a.진료년    = c.진료년                              " & vbCrLf
            SQL = SQL & "   And a.진료월    = c.진료월                              " & vbCrLf
            SQL = SQL & "   And a.진료일    = c.진료일                              " & vbCrLf
            SQL = SQL & "   And a.챠트번호  = c.챠트번호                            " & vbCrLf
            SQL = SQL & "   And (a.검사결과 IS NULL OR a.검사결과 = '')             " & vbCrLf
            SQL = SQL & " Order By ITEM                                             " & vbCrLf

'            SQL = ""
'            SQL = SQL & "Select DISTINCT (a.처방코드 + a.서브코드)      AS ITEM     " & vbCrlf
'            SQL = SQL & "  from tb_검사항목 " & vbCrlf
'            SQL = SQL & " Where 챠트번호 = '" & argPID & "'" & vbCrlf
'            SQL = SQL & "   And 진료년   = '" & strYear & "'" & vbCrlf
'            SQL = SQL & "   And 진료월   = '" & strMonth & "'" & vbCrlf
'            SQL = SQL & "   And 진료일   = '" & strDay & "'" & vbCrlf
'            SQL = SQL & "   And 처방번호 > 0 " & vbCrlf
'            SQL = SQL & "   And (검사결과 is null or 검사결과 = '') " & vbCrlf
'            SQL = SQL & "   And 처방코드+서브코드 in (" & gAllExam & ")"
        
        Case "MEDITOLISS"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT B.EXAM_CODE  AS ITEM                           " & vbCrLf
            SQL = SQL & "  FROM MEDITOLISS..TOTAL A, MEDITOLISS..TOTRES B               " & vbCrLf
            SQL = SQL + " WHERE A.EXAM_NO       = '" & strBarcode & "'                  " & vbCrLf
            SQL = SQL & "   And B.EXAM_CODE     IN (" & gAllTestCd & ")                 " & vbCrLf
            SQL = SQL & "   AND B.EXAM_PART     = 'C'                                   " & vbCrLf
            SQL = SQL & "   AND B.RESULT_VALUE  = ''                                    " & vbCrLf
            SQL = SQL & "   AND A.REQUEST_DATE  = B.REQUEST_DATE                        " & vbCrLf
            SQL = SQL & "   AND A.EXAM_NO       = B.EXAM_NO                             " & vbCrLf
                    
        Case "MOD"
            SQL = ""
            SQL = SQL & "Select Distinct c.EXAMCODE   AS ITEM           " & vbCrLf
            SQL = SQL & "  From EXAMREQ a, EXAMRES c                    " & vbCrLf
            SQL = SQL & " Where a.PID           = c.PID                 " & vbCrLf
            SQL = SQL & "   And a.SEQNO         = c.SEQNO               " & vbCrLf
            SQL = SQL & "   And a.RECENO        = c.RECENO              " & vbCrLf
            SQL = SQL & "   And c.SPECIMENID    = '" & strBarcode & "'  " & vbCrLf
            SQL = SQL & "   And c.EXAMCODE in (" & gAllTestCd & ")      " & vbCrLf
            SQL = SQL & "   And (c.EXAMEND = '' Or c.EXAMEND IS NULL)   " & vbCrLf
            SQL = SQL & " Order By c.EXAMCODE                           " & vbCrLf
    
        Case "MSINFOTEC"
            SQL = ""
            SQL = SQL & "Select DISTINCT ORCD AS ITEM       " & vbCrLf
            SQL = SQL & "  From emr.LRESULT                 " & vbCrLf
            SQL = SQL & " Where PAID = '" & strPatID & "'   " & vbCrLf
            SQL = SQL & "   And SPNO =  '" & strBarcode & "'" & vbCrLf
            SQL = SQL & "   And ORCD IN (" & gAllTestCd & ")" & vbCrLf
            SQL = SQL & "   And OKFL <> 'Y'                 " & vbCrLf   '-- 결과확정유무
            SQL = SQL & " Order By ORCD                     " & vbCrLf
        
        Case "NEOSOFT"
            If strInOut = "입원" Then
                SQL = ""
                SQL = SQL & "SELECT DISTINCT a.CODE as ITEM                         " & vbCr
                SQL = SQL & "  From E_ORDER..ORDER_IN" & Format(Now, "yyyy") & " a  " & vbCr
                SQL = SQL & " Where a.CHAM_INDEX =  '" & strBarcode & "'            " & vbCr
                SQL = SQL & "   AND a.CODE IN (" & gAllTestCd & ")                  " & vbCr
                SQL = SQL & "   AND a.TRANS = '2'                                   " & vbCr
                SQL = SQL & " ORDER BY a.CODE                                       " & vbCr
            ElseIf strInOut = "외래" Then
                SQL = ""
                SQL = SQL & "SELECT DISTINCT a.CODE as ITEM                         " & vbCr
                SQL = SQL & "  From E_ORDER..ORDER_OUT" & Format(Now, "yyyy") & " a " & vbCr
                SQL = SQL & " Where a.CHAM_INDEX =  '" & strBarcode & "'            " & vbCr
                SQL = SQL & "   AND a.CODE IN (" & gAllTestCd & ")                  " & vbCr
                SQL = SQL & "   AND a.TRANS = '2'                                   " & vbCr
                SQL = SQL & " ORDER BY a.CODE                                       " & vbCr
            End If
        
        Case "ONITGUM"
            SQL = ""
            SQL = SQL & "SELECT EDPSCODE     AS ITEM              " & vbCr
            SQL = SQL & "  FROM ONIT..GUMJIN_INTERFACE            " & vbCr
            SQL = SQL & " WHERE PER_GUM_NUM = '" & strBarcode & "'" & vbCr
            SQL = SQL & "   AND EDPSCODE IN (" & gAllTestCd & ")  " & vbCr
            SQL = SQL & "   AND (RESULT = ''  OR RESULT IS NULL)  " & vbCr
            
        Case "ONITEMR"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT b.MAP2SEQNO   AS ITEM      " & vbCr
            SQL = SQL & "  FROM " & gSQLDB.DB & "..WAITPRSNP a      " & vbCr
            SQL = SQL & "      ," & gSQLDB.DB & "..JUN370_RESULTTB b" & vbCr
            SQL = SQL & "      ," & gSQLDB.DB & "..PEWPRSNP c       " & vbCr
            SQL = SQL & "      ," & gSQLDB.DB & "..BAGMAP2PREF d    " & vbCr
            SQL = SQL & " WHERE a.WAITSEQNO = '" & strBarcode & "'  " & vbCr
            SQL = SQL & "   AND a.JUNDAL    = '" & gHOSP.HOSPCD & "'" & vbCr        '370
            SQL = SQL & "   AND a.WAITSEQNO = b.WAITSEQNO           " & vbCr
            SQL = SQL & "   AND a.CHARTNO   = c.CHARTNO             " & vbCr
            SQL = SQL & "   AND d.LABNO     IN (" & gHOSP.LABCD & ")" & vbCr   '4
            SQL = SQL & "   AND b.MAP2SEQNO IN (" & gAllTestCd & ") " & vbCr
            SQL = SQL & "   AND b.MAP2SEQNO = d.MAP2SEQNO           " & vbCr
            SQL = SQL & "   AND (b.RESULT = '' OR b.RESULT IS NULL) " & vbCr
        
        Case "PLIS"
            If Len(strBarcode) >= 11 Then
                strSpcYY = Mid(strBarcode, 1, 2)
                strSpcNo = Mid(strBarcode, 3, 9)
            Else
                Exit Function
            End If
            
            SQL = ""
            SQL = SQL & "SELECT DISTINCT r.testcd AS ITEM        " & vbCr
            SQL = SQL & "  FROM plis..s2lab201 m                 " & vbCr
            SQL = SQL & "     , plis..s2lab302 r                 " & vbCr
            SQL = SQL & "     , plis..s2lab001 e                 " & vbCr
            SQL = SQL & " WHERE m.spcyy = '" & strSpcYY & "'     " & vbCr
            SQL = SQL & "   and m.spcno = '" & strSpcNo & "'     " & vbCr
            SQL = SQL & "   and r.testcd IN (" & gAllTestCd & ") " & vbCr
            SQL = SQL & "   and (r.vfydt IS NULL OR r.vfydt='')  " & vbCr
            SQL = SQL & "   and m.workarea = r.workarea          " & vbCr
            SQL = SQL & "   and m.accdt = r.accdt                " & vbCr
            SQL = SQL & "   and m.accseq = r.accseq              " & vbCr
            SQL = SQL & "   and r.testcd = e.testcd              " & vbCr
            SQL = SQL & "  Order by r.testcd                     " & vbCr
        
        Case "TWIN"
            SQL = ""
            'SQL = SQL & "SELECT DISTINCT A.MASTERCODE AS ITEM           " & vbCr
            SQL = SQL & "SELECT DISTINCT A.SUBCODE    AS ITEM           " & vbCr
            SQL = SQL & "  From TW_HSP_OCS.TWEXAM_RESULTC A             " & vbCr
            SQL = SQL & "     , TW_HSP_OCS.TWEXAM_MASTER  B             " & vbCr
            SQL = SQL & "     , TW_HSP_OCS.TWEXAM_SPECMST C             " & vbCr
            SQL = SQL & " Where A.SPECNO =  '" & strBarcode & "'        " & vbCr
            SQL = SQL & "   And B.EQUCODE1 = '" & gHOSP.MACHCD & "'     " & vbCr '장비코드
            SQL = SQL & "   AND A.MASTERCODE IN (" & gAllTestCd & ")    " & vbCr
            SQL = SQL & "   AND C.STATUS   <= '3'                       " & vbCr '검사상태
            SQL = SQL & "   And C.SPECNO  = A.SPECNO                    " & vbCr
            SQL = SQL & "   And A.MASTERCODE = B.MASTERCODE             " & vbCr
            SQL = SQL & " ORDER BY A.ITEM                               " & vbCr
                
        Case "UBCARE"
            SQL = ""
            SQL = SQL & "Select Distinct EXAMCODE AS ITEM       " & vbCr
            SQL = SQL & "  From UB_PATRESULT                    " & vbCr
            SQL = SQL & " Where BARCODE = '" & strBarcode & "'  " & vbCr
            SQL = SQL & "   And EXAMCODE IN (" & gAllTestCd & ")" & vbCr
            SQL = SQL & "   And (RESULT = '' OR RESULT IS NULL) " & vbCr
            SQL = SQL & " Order by EXAMCODE                     " & vbCr
        
            Call SetSQLData("ITEM조회", SQL)
    
            '-- Record Count 가져옴
            AdoCn_Local.CursorLocation = adUseClient
            Set RS = AdoCn_Local.Execute(SQL, , 1)
            If Not RS.EOF = True And Not RS.BOF = True Then
                Do Until RS.EOF
                    If strItems = "" Then
                        strItems = GetTestNm(Trim(RS.Fields("ITEM")) & "", False)
                    Else
                        strItems = strItems & "/" & GetTestNm(Trim(RS.Fields("ITEM")), False)
                    End If
                    RS.MoveNext
                Loop
            End If
            
            GetSampleITEM = strItems
            
            RS.Close
            
            Exit Function
            
    End Select
            
                
    Call SetSQLData("ITEM조회", SQL)
    
    If SQL <> "" Then
        
        gPatOrdCd = ""
        
        '-- Record Count 가져옴
        AdoCn.CursorLocation = adUseClient
        Set RS = AdoCn.Execute(SQL, , 1)
        If Not RS.EOF = True And Not RS.BOF = True Then
            Do Until RS.EOF
                If strItems = "" Then
                    strItems = GetTestNm(Trim(RS.Fields("ITEM")) & "", False)
                Else
                    strItems = strItems & "/" & GetTestNm(Trim(RS.Fields("ITEM")), False)
                End If
                
                '처방 검사정보를 가져온다.
                gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("ITEM")) & "',"
                
                RS.MoveNext
            Loop
        End If
        
        GetSampleITEM = strItems
        
        RS.Close
    Else
        GetSampleITEM = ""
    End If
    
End Function


Public Sub LetEqpMaster(ByVal pEqpCD As String)

    SQL = ""
    SQL = SQL & "UPDATE EQPMASTER SET EQUIPCD = '" & pEqpCD & "'"

    Call DBExec(AdoCn_Local, SQL)

End Sub

'-- 장비결과 조회
Public Sub GetResultList(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As Object)
    Dim RS          As ADODB.Recordset
    Dim i           As Integer
    Dim iCnt        As Long
    Dim intRow      As Long
    Dim intCol      As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim strBarcode  As String
    Dim blnSame     As Boolean
    Dim strItems    As String
    Dim intOCnt     As Integer
    Dim strSaveSeq  As String
    Dim strExamDate As String

    Screen.MousePointer = 11
    iCnt = 0

    SQL = ""
    SQL = SQL & "SELECT DISTINCT SAVESEQ,EXAMDATE,HOSPDATE,EQUIPNO,BARCODE,SAMPLETYPE,DISKNO,POSNO" & vbCrLf
    SQL = SQL & ",CHARTNO,INOUT,PID,PNAME,PSEX,PAGE,PJUMIN,SENDFLAG,SENDDATE,EXAMUID,HOSPITAL " & vbCrLf
    SQL = SQL & ",SEQNO,EXAMNAME,RESULT,REFJUDGE" & vbCrLf
    SQL = SQL & "  FROM UB_PATRESULT " & vbCrLf
    SQL = SQL & " WHERE HOSPDATE Between '" & pFrom & "' AND '" & pTo & "'" & vbCrLf
    SQL = SQL & "   AND EXAMCODE IN (" & gAllTestCd & ") " & vbCrLf
    SQL = SQL & " ORDER BY HOSPDATE,BARCODE,SEQNO,SAVESEQ "

    '-- Record Count 가져옴
    AdoCn_Local.CursorLocation = adUseClient
    Set RS = AdoCn_Local.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        strItems = ""
        Do Until RS.EOF
            With SPD
                .ReDraw = False

                strSaveSeq = GetText(SPD, intRow, colSAVESEQ)
                strExamDate = GetText(SPD, intRow, colEXAMDATE)
                strBarcode = GetText(SPD, intRow, colBARCODE)
                iCnt = iCnt + 1

                If strSaveSeq <> Trim(RS.Fields("SAVESEQ")) & "" Or strBarcode <> Trim(RS.Fields("BARCODE")) & "" Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows

                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("SAVESEQ")) & "", intRow, colSAVESEQ
                    SetText SPD, Trim(RS.Fields("EXAMDATE")) & "", intRow, colEXAMDATE
                    'SetText SPD, Trim(RS.Fields("EXAMTIME")) & "", intRow, colEXAMTIME
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("DISKNO")) & "", intRow, colRACKNO
                    SetText SPD, Trim(RS.Fields("POSNO")) & "", intRow, colPOSNO
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("PSEX")) & "", intRow, colPSEX
                    SetText SPD, Trim(RS.Fields("PAGE")) & "", intRow, colPAGE
                    'SetText SPD, CStr(iCnt), intRow, colRCNT


                    Select Case Trim(RS.Fields("SENDFLAG")) & ""
                    Case "0"
                            SetText SPD, "장비결과", intRow, colSTATE
                    Case "1"
                            SetText SPD, "저장에러", intRow, colSTATE
                    Case "2"
                            SetText SPD, "전송완료", intRow, colSTATE
                    End Select

                End If

                For intCol = colSTATE + 1 To .MaxCols
                    .Row = 0
                    .Col = intCol
                    If Trim(RS.Fields("EXAMNAME")) & "" = Trim(.Text) Then
                        SetText SPD, Trim(RS.Fields("RESULT")) & "", intRow, intCol
                        If Trim(RS.Fields("REFJUDGE")) & "" <> "" Then
                            SetForeColor SPD, intRow, intRow, intCol, intCol, 255, 0, 0
                        End If
                        Exit For
                    End If

                Next

            End With
            DoEvents

            RS.MoveNext
        Loop
        'frmInterface.chkRAll.Value = "1"
    Else
        'frmInterface.lblStatus.Caption = ">> 조회 대상자가 없습니다."
        'frmInterface.chkRAll.Value = "0"
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

'    Call frmInterface.GetPatTRestResult_Search(1)

    Screen.MousePointer = 0

End Sub

'
''-- 장비결과 조회
''Public Sub GetResultList(ByVal pFrom As String, ByVal pTo As String, ByVal pDateType As Integer, ByVal pOpt As Integer)
'Public Sub GetResultList(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As Object)
'    Dim RS          As ADODB.Recordset
'    Dim i           As Integer
'    Dim iCnt        As Long
'    Dim intRow      As Long
'    Dim intCol      As Integer
'    Dim strDate     As String
'    Dim strChart    As String
'    Dim strBarcode  As String
'    Dim blnSame     As Boolean
'    Dim strItems    As String
'    Dim intOCnt     As Integer
'    Dim strSaveSeq  As String
'    Dim strExamDate As String
'
'    Screen.MousePointer = 11
'    iCnt = 0
'
'    SQL = ""
'    SQL = SQL & "SELECT DISTINCT SAVESEQ,EXAMDATE,HOSPDATE,EQUIPNO,BARCODE,SAMPLETYPE,DISKNO,POSNO" & vbCr
'    SQL = SQL & ",CHARTNO,INOUT,PID,PNAME,PSEX,PAGE,PJUMIN,SENDFLAG,SENDDATE,EXAMUID,HOSPITAL " & vbCr
'    '-- 검사결과
'    SQL = SQL & ",SEQNO,EXAMNAME,RESULT,REFJUDGE" & vbCr
'
'    If gEMR = "UBCARE" Then
'        SQL = SQL & "  FROM UB_PATRESULT " & vbCr
'    Else
'        SQL = SQL & "  FROM PATRESULT " & vbCr
'    End If
'
'    '-- 검사일자
'    'If pDateType = 0 Then
'        SQL = SQL & " WHERE EXAMDATE Between '" & pFrom & "' AND '" & pTo & "'" & vbCr
'    '-- 접수일자
'    'Else
'    '    SQL = SQL & " WHERE HOSPDATE Between '" & pFrom & "' AND '" & pTo & "'" & vbCr
'    'End If
'
'    '-- 전송
'    'If pOpt = 1 Then
'    '    SQL = SQL & "   AND SENDFLAG = '2' " & vbCr
'    '-- 미전송
'    'ElseIf pOpt = 2 Then
'    '    SQL = SQL & "   AND SENDFLAG <> '2' " & vbCr
'    'End If
'
'    SQL = SQL & "   AND EXAMCODE IN (" & gAllTestCd & ") " & vbCr
'
'    'If pDateType = 0 Then
'    '    SQL = SQL & " ORDER BY EXAMDATE,SAVESEQ,BARCODE,SEQNO"
'    'Else
'        SQL = SQL & " ORDER BY HOSPDATE,BARCODE,SEQNO,SAVESEQ "
'    'End If
'
'    '-- Record Count 가져옴
'    AdoCn_Local.CursorLocation = adUseClient
'    Set RS = AdoCn_Local.Execute(SQL, , 1)
'    If Not RS.EOF = True And Not RS.BOF = True Then
'        SPD.MaxRows = 0
'        strItems = ""
'        Do Until RS.EOF
'            With SPD
'            iCnt = iCnt + 1
'            With SPD
'                .ReDraw = False
'
''                If iCnt = 1 Then
''                    .MaxRows = .MaxRows + 1
''                    intRow = .MaxRows
''                End If
'
'                strSaveSeq = GetText(SPD, intRow, colSAVESEQ)
'                strBarcode = GetText(SPD, intRow, colBARCODE)
'                strExamDate = GetText(SPD, intRow, colEXAMDATE)
'
'                'If strSaveSeq <> Trim(RS.Fields("SAVESEQ")) & "" Or strExamDate <> Trim(RS.Fields("EXAMDATE")) & "" Then
'                If strBarcode <> Trim(RS.Fields("BARCODE")) & "" Or strExamDate <> Trim(RS.Fields("EXAMDATE")) & "" Then
'                    .MaxRows = .MaxRows + 1
'                    intRow = .MaxRows
'
'                    SetText SPD, "1", intRow, colCHECKBOX
'                    SetText SPD, Trim(RS.Fields("SAVESEQ")) & "", intRow, colSAVESEQ
'                    SetText SPD, Trim(RS.Fields("EXAMDATE")) & "", intRow, colEXAMDATE
'                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
'                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
'                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
'                    SetText SPD, Trim(RS.Fields("DISKNO")) & "", intRow, colRACKNO
'                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
'                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
'                    SetText SPD, Trim(RS.Fields("PSEX")) & "", intRow, colPSEX
'                    SetText SPD, Trim(RS.Fields("PAGE")) & "", intRow, colPAGE
'                    SetText SPD, Trim(RS.Fields("PJUMIN")) & "", intRow, colPJUMIN
'                    SetText SPD, Trim(RS.Fields("INOUT")) & "", intRow, colINOUT
''                    SetText SPD, Trim(RS.Fields("EQUIPNO")) & "", intRow, colKEY1
'
'
'                    Select Case Trim(RS.Fields("SENDFLAG")) & ""
'                    Case "0"
'                            SetText SPD, "장비결과", intRow, colSTATE
'                    Case "2"
'                            SetText SPD, "전송완료", intRow, colSTATE
'                    End Select
'
'                    If gEMR <> "KOMAIN" Then
'                        'SetText spd, GetSampleITEM(intRow, spd), intRow, colITEMS
'                    End If
'                End If
'
'                For intCol = colSTATE + 1 To .MaxCols
'                    .Row = 0
'                    .Col = intCol
'                    If Trim(RS.Fields("EXAMNAME")) & "" = Trim(.Text) Then
'                        SetText SPD, Trim(RS.Fields("RESULT")) & "", intRow, intCol
'                        If Trim(RS.Fields("REFJUDGE")) & "" <> "" Then
'                            SetForeColor SPD, intRow, intRow, intCol, intCol, 255, 0, 0
'                        End If
'                        Exit For
'                    End If
'
'                Next
'
'            End With
'            DoEvents
'
'            RS.MoveNext
'            End With
'        Loop
'        '.chkRAll.Value = "1"
'    Else
'        'frmMain.lblStatus.Caption = ">> 조회 대상자가 없습니다."
'        'frmMain.chkRAll.Value = "0"
'    End If
'
'    RS.Close
'
'    SPD.RowHeight(-1) = 15
'    SPD.ReDraw = True
'
'    'Call frmMain.GetPatTRestResult_Search(1)
'
'    Screen.MousePointer = 0
'
'End Sub

'-- 검사결과 서버저장
Function SaveTransData(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    
    SaveTransData = -1
    
    Select Case gEMR
        Case "UBCARE"
            SaveTransData = SaveTransData_UBCARE(argSpcRow, SPD)
        
        Case "BIT"
            SaveTransData = SaveTransData_BIT(argSpcRow, SPD)
        
        Case "EASYS"
            SaveTransData = SaveTransData_EASYS(argSpcRow, SPD)
        
        Case "HEALTH"
'            SaveTransData = SaveTransData_HEALTH(argSpcRow, SPD)
        
        Case "BITJSON"
            SaveTransData = SaveTransData_BITJSON(argSpcRow, SPD)
        
        Case "NEOSENSE"
            SaveTransData = SaveTransData_NEOSENSE(argSpcRow, SPD)
        
        Case "MCC"
            SaveTransData = SaveTransData_MCC(argSpcRow, SPD)
        
        Case "JWINFO"
            SaveTransData = SaveTransData_JWINFO(argSpcRow, SPD)

        Case "SCL"
            SaveTransData = SaveTransData_SCL(argSpcRow, SPD)
        
        Case "KCHART"
            SaveTransData = SaveTransData_KCHART(argSpcRow, SPD)
        
        Case "AMIS"
            SaveTransData = SaveTransData_AMIS(argSpcRow, SPD)
        
        Case "MEDICHART"
            SaveTransData = SaveTransData_MEDICHART(argSpcRow, SPD)
        
        Case "LABSPEAR"
            SaveTransData = SaveTransData_LABSPEAR(argSpcRow, SPD)
        
        Case "SANSOFT"
            SaveTransData = SaveTransData_LABSPEAR(argSpcRow, SPD)
        
        Case "BIT70"
            SaveTransData = SaveTransData_BIT70(argSpcRow, SPD)
        
        Case "EONM"
            SaveTransData = SaveTransData_EONM(argSpcRow, SPD)
        
        Case "TWIN"
            SaveTransData = SaveTransData_TWIN(argSpcRow, SPD)
    End Select


End Function

'-- 검사결과 서버저장
Function SaveTransDataR(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    
    SaveTransDataR = -1
    
    Select Case gEMR
        Case "AMIS"
'            SaveTransDataR = SaveTransDataR_AMIS(argSpcRow)
        
'        Case "BIGUBCARE"
'            SaveTransDataR = SaveTransDataR_BIGUBCARE(argSpcRow)
'
'        Case "BIT"
'            SaveTransDataR = SaveTransDataR_BIT(argSpcRow)
'
'        Case "BIT70"
'            SaveTransDataR = SaveTransDataR_BIT70(argSpcRow)
'
'        Case "EMEDI"
'            SaveTransDataR = SaveTransDataR_AMIS(argSpcRow)
'
'        Case "EONM"
'            SaveTransDataR = SaveTransDataR_EONM(argSpcRow)
'
'        Case "EASYS"
'            SaveTransDataR = SaveTransDataR_EASYS(argSpcRow)
'
'        Case "GINUS"
'            SaveTransDataR = SaveTransDataR_GINUS(argSpcRow)
'
'        Case "GSEN"
'            SaveTransDataR = SaveTransDataR_MSINFOTEC(argSpcRow)
'
'        Case "HWASAN"
'            SaveTransDataR = SaveTransDataR_HWASAN(argSpcRow)
'
'        Case "JAINCOM"
'            SaveTransDataR = SaveTransDataR_JAINCOM(argSpcRow)
'
'        Case "JWINFO"
'            SaveTransDataR = SaveTransDataR_JWINFO(argSpcRow)
'
'        Case "KCHART"
'            SaveTransDataR = SaveTransDataR_KCHART(argSpcRow)
'
'        Case "KOMAIN"
'            SaveTransDataR = SaveTransDataR_KOMAIN(argSpcRow)
'
'        Case "KYU"
'            SaveTransDataR = SaveTransDataR_KYU(argSpcRow)
'
'        Case "MEDICHART"
'            SaveTransDataR = SaveTransDataR_MEDICHART(argSpcRow)
'
'        Case "MEDIIT"
'            SaveTransDataR = SaveTransDataR_MEDIIT(argSpcRow)
'
'        Case "MEDITOLISS"
'            SaveTransDataR = SaveTransDataR_MEDITOLISS(argSpcRow)
'
'        Case "MCC"
'            SaveTransDataR = SaveTransDataR_MCC(argSpcRow)
'
'        Case "MOD"
'            SaveTransDataR = SaveTransDataR_MOD(argSpcRow)
'
'        Case "MSINFOTEC"
'            SaveTransDataR = SaveTransDataR_MSINFOTEC(argSpcRow)
'
'        Case "NEOSOFT"
'            SaveTransDataR = SaveTransDataR_NEOSOFT(argSpcRow)
'
'        Case "ONITGUM"
'            SaveTransDataR = SaveTransDataR_ONITGUM(argSpcRow)
'
'        Case "ONITEMR"
'            SaveTransDataR = SaveTransDataR_ONITEMR(argSpcRow)
'
'        Case "PLIS"
'            SaveTransDataR = SaveTransDataR_PLIS(argSpcRow)
'
'        Case "SY"
'            SaveTransDataR = SaveTransDataR_SY(argSpcRow)
'
'        Case "TWIN"
'            SaveTransDataR = SaveTransDataR_TWIN(argSpcRow)
'
'        Case "UBCARE"
'            SaveTransDataR = SaveTransDataR_UBCARE(argSpcRow)

        
        Case Else
            SaveTransDataR = -1
    End Select

End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 1. 접수정보 조회,
'                                 2. 장비수신정보 화면표시,
'                                 3. 처방코드 가져오기
'   인수 :
'       - pBarNo : 바코드번호
'       - pType  : 바코드 미사용시 비교하는 대상
'                   1 : Seq
'                   2 : Rack/Pos
'                   3 : 체크된것중 제일 위에 것
'-----------------------------------------------------------------------------'
Public Sub SetPatInfo(ByVal pBarNo As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    'Call SetCommStatus("R", pBarNo, frmInterface.spdComStatus)
'    Call SetCommStatus("R", pBarNo, frmInterface.lstComStatus)
    
    intRow = -1
    With frmInterface
        Select Case pType
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmInterface.spdOrder, i, colBARCODE) = pBarNo Then
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmInterface.spdOrder, i, colSEQNO) = mOrder.Seq Then
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmInterface.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmInterface.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmInterface.spdOrder, i, colCHECKBOX) = "1" And GetText(frmInterface.spdOrder, i, colSTATE) = "" Then
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
        
        
        '검사순번이 같은지 찾는다.
'        If intRow > 0 Then
'            For i = 1 To .spdOrder.DataRowCnt
'                If GetText(frmInterface.spdOrder, i, colSAVESEQ) = mResult.RsltSeq Then
'                    intRow = i
'                    Exit For
'                End If
'            Next i
'        End If
        
        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
                .spdOrder.GridColor = &HE0E0E0
                .spdOrder.GridShowHoriz = True
                .spdOrder.GridShowVert = True
            End If
        End If
        
        '-- 장비결과인덱스 화면표시
        Call SetText(.spdOrder, "1", intRow, colCHECKBOX)
        Call SetText(.spdOrder, mResult.RsltDate, intRow, colEXAMDATE)
        Call SetText(.spdOrder, mResult.RsltTime, intRow, colEXAMTIME)
        Call SetText(.spdOrder, mResult.RsltSeq, intRow, colSAVESEQ)
        If pType = "0" Then
            Call SetText(.spdOrder, mResult.BarNo, intRow, colBARCODE)
        End If
'        Call SetText(.spdOrder, mResult.RackNo, intRow, colRACKNO)
'        Call SetText(.spdOrder, mResult.TubePos, intRow, colPOSNO)
'        Call SetText(.spdOrder, mResult.Seq, intRow, colSEQNO)
        
        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0
    
        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)
        
        .spdOrder.RowHeight(-1) = 15
        
    End With
    
    '-- 현재 Row
    gRow = intRow
    
End Sub

'-- 검사자 정보 가져오기
Function GetSampleInfo(ByVal asRow As Long, ByVal SPD As Object) As Integer

    Screen.MousePointer = 11

    GetSampleInfo = -1

    'If cn_Server_Flag = True Then
        Select Case gEMR
            Case "UBCARE"
                    Call GetSampleInfo_UBCARE(asRow, SPD)
            
            Case "EASYS"
                    Call GetSampleInfo_EASYS(asRow, SPD)
            
            Case "HEALTH"
'                    Call GetSampleInfo_HEALTH(asRow, SPD)
            
            Case "BITJSON"
                    Call GetSampleInfo_BITJSON(asRow, SPD)
            
            Case "SY"
                    Call GetSampleInfo_SY(asRow, SPD)
            
            Case "NEOSENSE"
                    Call GetSampleInfo_NEOSENSE(asRow, SPD)
            
            Case "MCC"
                    Call GetSampleInfo_MCC(asRow, SPD)
    
            Case "BIT"
                    Call GetSampleInfo_BIT(asRow, SPD)
            
            Case "JWINFO"
                    Call GetSampleInfo_JWINFO(asRow, SPD)
            
            Case "SCL"
                    Call GetSampleInfo_SCL(asRow, SPD)
    
            Case "KCHART"
                    Call GetSampleInfo_KCHART(asRow, SPD)
            
            Case "AMIS"
                    Call GetSampleInfo_AMIS(asRow, SPD)
            
            Case "MEDICHART"
                    Call GetSampleInfo_MEDICHART(asRow, SPD)
            
            Case "LABSPEAR"
                    Call GetSampleInfo_LABSPEAR(asRow, SPD)
            
            Case "SANSOFT"
                    Call GetSampleInfo_LABSPEAR(asRow, SPD)
            
            Case "BIT70"
                    Call GetSampleInfo_BIT70(asRow, SPD)
            
            Case "EONM"
                    Call GetSampleInfo_EONM(asRow, SPD)
            
            Case "TWIN"
                    Call GetSampleInfo_TWIN(asRow, SPD)
    
    
        End Select
    
        GetSampleInfo = 1
    
    'End If
    
    Screen.MousePointer = 0


End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_PHILL(ByVal asRow As Long, ByVal SPD As Object) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo            As Long
    
On Error GoTo ErrHandle
    
    GetSampleInfo_PHILL = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    strRegDate = Mid(strBarcode, 1, 8)
    lngRegNo = Val(Mid(strBarcode, 9))
    
    Screen.MousePointer = 11
          
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       P.request_date  AS HOSPDATE " & vbCrLf
    SQL = SQL & "     , P.exam_no       AS PID      " & vbCrLf
    SQL = SQL & "     , P.company_code  AS INOUT    " & vbCrLf
    SQL = SQL & "     , P.chart_no      AS CHARTNO  " & vbCrLf
    SQL = SQL & "     , p.personal_id               " & vbCrLf
    SQL = SQL & "     , p.person_name   AS PNAME    " & vbCrLf
    SQL = SQL & "     , P.worker_code               " & vbCrLf
    SQL = SQL & "     , P.patient_kind              " & vbCrLf
    SQL = SQL & "     , P.person_sex    AS SEX      " & vbCrLf
    SQL = SQL & "     , P.person_age    AS AGE      " & vbCrLf
    SQL = SQL & "     , R.exam_order                " & vbCrLf
    SQL = SQL & "     , R.exam_code     AS ITEM     " & vbCrLf
    SQL = SQL & "     , E.exam_ename                " & vbCrLf
    SQL = SQL & "     , R.pro_code      AS ORDERCODE            " & vbCrLf
    SQL = SQL & "  FROM trust P, trures R, examitem E           " & vbCrLf
    SQL = SQL & " WHERE P.request_date  = '" & strRegDate & "'  " & vbCrLf
    SQL = SQL & "   AND P.exam_no       = '" & lngRegNo & "'    " & vbCrLf
    SQL = SQL & "   AND R.exam_code     IN (" & gAllTestCd & ") " & vbCrLf
    SQL = SQL & "   AND R.exam_code     <> 'X999'               " & vbCrLf
    SQL = SQL & "   AND P.request_date  = R.request_date        " & vbCrLf
    SQL = SQL & "   AND P.exam_no       = R.exam_no             " & vbCrLf
    SQL = SQL & "   AND R.exam_code     = E.exam_code           " & vbCrLf
    SQL = SQL & " ORDER BY P.request_date, P.exam_no            "
        
    Call SetSQLData("바코드조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    
    SetText SPD, "0", asRow, colCHECKBOX
        
    ReDim Preserve gPatTest(RS.RecordCount)
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Format(Trim(RS.Fields("HOSPDATE")) & "", "####-##-##"), asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("patient_kind")) & "", asRow, colINOUT
                'SetText SPD, Trim(RS.Fields("BARCODE")), asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("CHARTNO")), asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '환자 성별/나이
                With mPatient
                    .AGE = Trim(RS.Fields("AGE")) & ""
                    .SEX = Trim(RS.Fields("SEX")) & ""
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        '-- 처방코드
                        gArrEQP(intCol - colSTATE, 16) = Trim(RS.Fields("ORDERCODE")) & ""
                        Exit For
                    End If
                Next
                
                gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("ITEM")) & "',"
                gPatTest(intTestCnt) = Trim(RS.Fields("ITEM"))
            End With
            DoEvents
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_PHILL = 1
    
    Screen.MousePointer = 0
    
Exit Function

ErrHandle:
    GetSampleInfo_PHILL = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "GetSampleInfo_PHILL" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function


'-- 검사자 정보 가져오기
Function GetSampleInfo_EONM(ByVal asRow As Long, ByVal SPD As Object) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo            As Long
    
On Error GoTo DBErr
    
    GetSampleInfo_EONM = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
        
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       O.H141_ODRDAT        AS HOSPDATE " & vbCrLf
    SQL = SQL & "      ,O.H141_TSAMPLENO     AS BARCODE  " & vbCrLf
    SQL = SQL & "      ,P.A110_CHARTNO       AS CHARTNO  " & vbCrLf
    SQL = SQL & "      ,P.A110_PATNM         AS PNAME    " & vbCrLf
    SQL = SQL & "      ,P.A110_JUMIN1        AS AGE      " & vbCrLf
    SQL = SQL & "      ,P.A110_SEX           AS SEX      " & vbCrLf
    SQL = SQL & "      ,O.H141_SUGACD        AS ITEM     " & vbCrLf
    SQL = SQL & "      ,O.H141_SEQNO         AS SUBITEM  " & vbCrLf
    SQL = SQL & "  FROM TB_H141_LISTAKEBODY O, TB_A110_PATINFO P  " & vbCrLf          'TB_H131_SPPRESULT
    SQL = SQL & " Where P.A110_ChartNo      = O.H141_CHARTNO      " & vbCrLf
    SQL = SQL & "   AND O.H141_TSAMPLENO    = '" & strBarcode & "'" & vbCrLf
    'SQL = SQL & "   AND O.H141_NOTYYN       = 'N'                 " & vbCr
    SQL = SQL & "   AND O.H141_NOTYYN       IN ('N','T')                 " & vbCr '결과대기:T
    SQL = SQL & "   And O.H141_SUGACD in (" & gAllTestCd & ")     " & vbCrLf
    SQL = SQL & " Order By O.H141_SUGACD                          " & vbCrLf
        
    Call SetSQLData("바코드조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    
    SetText SPD, "0", asRow, colCHECKBOX
        
    ReDim Preserve gPatTest(RS.RecordCount)
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("BARCODE")) & "", asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("CHARTNO")) & "", asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
'                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '환자 성별/나이
                With mPatient
                    .AGE = Trim(RS.Fields("AGE")) & ""
                    .SEX = Trim(RS.Fields("SEX")) & ""
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        
                        '-- 결과저장용 SEQ
                        gArrEQP(intCol - colSTATE, 16) = Trim(RS.Fields("SUBITEM")) & ""
                        
                        Exit For
                    End If
                Next
                
                gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("ITEM")) & "',"
                'gPatTest(intTestCnt) = Trim(RS.Fields("ITEM"))
            End With
            DoEvents
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_EONM = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_EONM = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "GetSampleInfo_EONM" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_AMIS(ByVal asRow As Long, ByVal SPD As Object) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo            As Long
    
On Error GoTo DBErr
    
    GetSampleInfo_AMIS = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
        
'    SQL = SQL & "   AND R.RESULTFLAG    = 0                     " & vbCrLf
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT"
    SQL = SQL & "       O.ACPTDATE              AS HOSPDATE             " & vbCrLf
    SQL = SQL & "     , R.SPCMNO                AS BARCODE              " & vbCrLf
    SQL = SQL & "     , P.PATID                 AS PID                  " & vbCrLf
    SQL = SQL & "     , P.PATNAME               AS PNAME                " & vbCrLf
    SQL = SQL & "     , P.SEX                   AS SEX                  " & vbCrLf
    SQL = SQL & "     , R.ORDERCODE             AS ORDCODE              " & vbCrLf
    SQL = SQL & "     , R.RESULTITEMCODE        AS ITEM                 " & vbCrLf
    SQL = SQL & "  FROM REGISTINFOS O                                   " & vbCrLf
    SQL = SQL & "     , RESULTOFNUM R                                   " & vbCrLf
    SQL = SQL & "     , PATMST      P                                   " & vbCrLf
    SQL = SQL & " WHERE O.ACPTDATE      = R.ACPTDATE                    " & vbCrLf
    SQL = SQL & "   AND O.PATID         = R.PATID                       " & vbCrLf
    SQL = SQL & "   AND O.ACPTSEQ       = R.ACPTSEQ                     " & vbCrLf
    SQL = SQL & "   AND O.PATID         = P.PATID                       " & vbCrLf
    SQL = SQL & "   AND O.CLAS          = 4                             " & vbCrLf '임상병리
    SQL = SQL & "   AND R.SPCMNO        = '" & strBarcode & "'          " & vbCrLf
    SQL = SQL & "   AND (R.NUMRESULTVAL = '' OR R.NUMRESULTVAL IS NULL) " & vbCrLf
    SQL = SQL & "   AND R.RESULTITEMCODE IN (" & gAllTestCd & ")        " & vbCrLf
        
    Call SetSQLData("바코드조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    
    SetText SPD, "0", asRow, colCHECKBOX
        
    ReDim Preserve gPatTest(RS.RecordCount)
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("BARCODE")) & "", asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '환자 성별/나이
                With mPatient
                    .SEX = Trim(RS.Fields("SEX")) & ""
                    '.AGE = Trim(RS.Fields("AGE")) & ""
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        
                        '-- 처방코드
                        gArrEQP(intCol - colSTATE, 16) = Trim(RS.Fields("ORDCODE")) & ""
                        
                        Exit For
                    End If
                Next
                
                gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("ITEM")) & "',"
                'gPatTest(intTestCnt) = Trim(RS.Fields("ITEM"))
            End With
            DoEvents
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_AMIS = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_AMIS = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetSampleInfo_AMIS" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_KCHART(ByVal asRow As Long, ByVal SPD As Object) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo            As Long
    
On Error GoTo DBErr
    
    GetSampleInfo_KCHART = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
        
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       J.접수일자                  AS HOSPDATE                 " & vbCrLf
    SQL = SQL & "     , L.검체번호                  AS BARCODE                  " & vbCrLf
    SQL = SQL & "     , A.챠트번호                  AS CHARTNO                  " & vbCrLf
    SQL = SQL & "     , J.접수번호                  AS PID                      " & vbCrLf
    SQL = SQL & "     , A.환자이름                  AS PNAME                    " & vbCrLf
    SQL = SQL & "     , A.환자성별                  AS SEX                      " & vbCrLf
    SQL = SQL & "     , A.환자나이                  AS AGE                      " & vbCrLf
    SQL = SQL & "     , L.진료검사ID                AS TESTID                   " & vbCrLf
    SQL = SQL & "     , L.진료지원ID                AS SPRTID                   " & vbCrLf
    SQL = SQL & "     , (L.처방코드+ L.서브코드)    AS ITEM                     " & vbCrLf
    SQL = SQL & "  FROM         TB_진료검사 L                                   " & vbCrLf
    SQL = SQL & "   INNER JOIN  TB_진료지원 J ON (L.진료지원ID = J.진료지원ID)  " & vbCrLf
    SQL = SQL & "   INNER JOIN  TB_진료일반 A ON (J.진료일자   = A.진료일자     " & vbCrLf
    SQL = SQL & "                            AND  J.챠트번호   = A.챠트번호     " & vbCrLf
    SQL = SQL & "                            AND  J.진료번호   = A.진료번호)    " & vbCrLf
    SQL = SQL & " Where L.검체번호 = '" & strBarcode & "'                       " & vbCrLf
    SQL = SQL & "   AND L.검사상태 < 5                                          " & vbCrLf
    SQL = SQL & "   AND L.처방코드 + L.서브코드 IN (" & gAllTestCd & ")         " & vbCrLf
    SQL = SQL & " ORDER BY J.접수일자, L.검체번호                               " & vbCrLf
    Call SetSQLData("바코드조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    
    SetText SPD, "0", asRow, colCHECKBOX
        
    ReDim Preserve gPatTest(RS.RecordCount)
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("BARCODE")) & "", asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("CHARTNO")) & "", asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '환자 성별/나이
                With mPatient
                    .SEX = Trim(RS.Fields("SEX")) & ""
                    .AGE = Trim(RS.Fields("AGE")) & ""
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        
                        '-- 진료검사ID
                        gArrEQP(intCol - colSTATE, 16) = Trim(RS.Fields("TESTID")) & ""
                        
                        '-- 진료지원ID
                        gArrEQP(intCol - colSTATE, 17) = Trim(RS.Fields("SPRTID")) & ""
                        
                        Exit For
                    End If
                Next
                
                gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("ITEM")) & "',"
            End With
            DoEvents
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_KCHART = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_KCHART = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetSampleInfo_KCHART" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_JWINFO(ByVal asRow As Long, ByVal SPD As Object) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo            As Long
    
On Error GoTo DBErr
    
    GetSampleInfo_JWINFO = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
        
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       a.RECEIPTDATE           AS HOSPDATE " & vbCrLf
    SQL = SQL & "     , a.SPECIMENNUM           AS BARCODE  " & vbCrLf
    SQL = SQL & "     , a.RECEIPTNO             AS CHARTNO  " & vbCrLf
    SQL = SQL & "     , a.IPDOPD                AS INOUT    " & vbCrLf
    SQL = SQL & "     , a.PTNO                  AS PID      " & vbCrLf
    SQL = SQL & "     , a.SNAME                 AS PNAME    " & vbCrLf
    SQL = SQL & "     , a.ORDERCODE             AS ORDCODE  " & vbCrLf
    SQL = SQL & "     , b.LABCODE               AS ITEM     " & vbCrLf
    SQL = SQL & "   FROM SLA_LabMaster a, SLA_LabResult b   " & vbCrLf
    SQL = SQL & " WHERE a.RECEIPTNO     = b.RECEIPTNO       " & vbCrLf
    SQL = SQL & "   AND a.ORDERCODE     = b.ORDERCODE       " & vbCrLf
    SQL = SQL & "   and a.RECEIPTDATE   = b.RECEIPTDATE     " & vbCrLf
    SQL = SQL & "   AND a.SPECIMENNUM   = b.SPECIMENNUM     " & vbCrLf
    SQL = SQL & "   AND a.SPECIMENNUM   = '" & strBarcode & "'" & vbCrLf
    SQL = SQL & "   AND b.LABCODE IN (" & gAllTestCd & ")   " & vbCrLf
'    SQL = SQL & "   AND a.JSTATUS < '3'                     " & vbCrLF
    SQL = SQL & " ORDER BY b.LABCODE                        " & vbCrLf
    
    Call SetSQLData("바코드조회", SQL)
    
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    
    SetText SPD, "0", asRow, colCHECKBOX
        
    ReDim Preserve gPatTest(RS.RecordCount)
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("BARCODE")) & "", asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("CHARTNO")) & "", asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        
                        '-- 처방코드
                        gArrEQP(intCol - colSTATE, 16) = Trim(RS.Fields("ORDCODE")) & ""
                        
                        Exit For
                    End If
                Next
                
                gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("ITEM")) & "',"
            End With
            DoEvents
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_JWINFO = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_JWINFO = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetSampleInfo_JWINFO" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_SCL(ByVal asRow As Long, ByVal SPD As Object) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo        As Long
    
    Dim Prm1            As New ADODB.Parameter
    Dim Prm2            As New ADODB.Parameter
    Dim Prm3            As New ADODB.Parameter
    
On Error GoTo DBErr
    
    GetSampleInfo_SCL = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
        
    '-- SP 사용
    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn

    AdoCmd.CommandTimeout = 15
    AdoCmd.CommandText = "SLRTRM51P"
    AdoCmd.CommandType = adCmdStoredProc
    
    Set Prm1 = AdoCmd.CreateParameter("pbarc", adChar, adParamInput, 12, strBarcode)
    AdoCmd.Parameters.Append Prm1
    
    'Set Prm2 = AdoCmd.CreateParameter("pmach", adChar, adParamInput, 3, gHOSP.MACHCD)
    Set Prm2 = AdoCmd.CreateParameter("pmach", adChar, adParamInput, 3, mResult.EqpCd)
    AdoCmd.Parameters.Append Prm2
        
    Set Prm3 = AdoCmd.CreateParameter("perr", adChar, adParamOutput, 1, "")
    AdoCmd.Parameters.Append Prm3
    
    Set RS = New ADODB.Recordset
    RS.Open AdoCmd.Execute
    
    SetText SPD, "0", asRow, colCHECKBOX
        
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(RS.Fields("ORDDATE")) & "", asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("BARCODENO")) & "", asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("WORKNO")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODENO")) & ""
                    .PID = Trim(RS.Fields("WORKNO")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEMCODE")) & Trim(RS.Fields("DCODE")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        
                        Exit For
                    End If
                Next
                
                gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("ITEMCODE")) & Trim(RS.Fields("DCODE")) & "',"
            End With
            DoEvents
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_SCL = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_SCL = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetSampleInfo_SCL" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

Function GetSampleInfo_MEDICHART(ByVal asRow As Long, ByVal SPD As Object) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo            As Long
    
On Error GoTo DBErr
    
    GetSampleInfo_MEDICHART = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
        
    SQL = ""
    SQL = SQL & "Select DISTINCT "
    SQL = SQL & "       (a.진료년 + a.진료월 + a.진료일)    AS HOSPDATE     " & vbCrLf
    SQL = SQL & "     , a.챠트번호                          AS CHARTNO      " & vbCrLf
    SQL = SQL & "     , c.진료상태                          AS STATE        " & vbCrLf
    SQL = SQL & "     , b.수진자명                          AS PNAME        " & vbCrLf
    SQL = SQL & "     , b.주민등록번호                      AS PJUMIN       " & vbCrLf
    SQL = SQL & "     , (a.처방코드 + a.서브코드)           AS ITEM         " & vbCrLf
    SQL = SQL & "  From TB_검사항목 a, TB_인적사항 b, TB_진료기본 c         " & vbCrLf
    SQL = SQL & " Where a.챠트번호 = '" & strChartNo & "'                   " & vbCrLf
    SQL = SQL & "   And a.처방번호 > 0                                      " & vbCrLf
    SQL = SQL & "   And c.진료상태 IN ('1','5','6','7','8','9')             " & vbCrLf
'    SQL = SQL & "   And (a.처방코드 + a.서브코드) IN (" & gAllTestCd & ")   " & vbCrLf
    SQL = SQL & "   And (a.처방코드 + '|' + a.서브코드) IN (" & gAllTestCd & ")   " & vbCrLf
    SQL = SQL & "   And (a.검사결과 IS NULL OR a.검사결과 = '')             " & vbCrLf
    SQL = SQL & "   And a.진료년    = c.진료년                              " & vbCrLf
    SQL = SQL & "   And a.진료월    = c.진료월                              " & vbCrLf
    SQL = SQL & "   And a.진료일    = c.진료일                              " & vbCrLf
    SQL = SQL & "   And a.챠트번호  = c.챠트번호                            " & vbCrLf
    SQL = SQL & "   And a.챠트번호  = b.챠트번호                            " & vbCrLf
    SQL = SQL & "   And (a.검사결과 IS NULL OR a.검사결과 = '')             " & vbCrLf
    SQL = SQL & " Order By a.진료년, a.진료월, a.진료일, b.수진자명         " & vbCrLf
        
    Call SetSQLData("바코드조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    
    SetText SPD, "0", asRow, colCHECKBOX
        
    ReDim Preserve gPatTest(RS.RecordCount)
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("CHARTNO")) & "", asRow, colCHARTNO
                'SetText SPD, Trim(RS.Fields("BARCODE")) & "", asRow, colBARCODE
                'SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
               ' SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '환자 성별/나이
                With mPatient
                    '.SEX = Trim(RS.Fields("SEX")) & ""
                    '.AGE = Trim(RS.Fields("AGE")) & ""
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        
                        '-- 처방코드
'                        gArrEQP(intCol - colSTATE, 16) = Trim(RS.Fields("ORDCODE")) & ""
                        
                        Exit For
                    End If
                Next
                
                gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("ITEM")) & "',"
                'gPatTest(intTestCnt) = Trim(RS.Fields("ITEM"))
            End With
            DoEvents
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_MEDICHART = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_MEDICHART = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetSampleInfo_MEDICHART" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_LABSPEAR(ByVal asRow As Long, ByVal SPD As Object) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo            As Long
    
    Dim strRegno            As String
    
On Error GoTo DBErr
    
    GetSampleInfo_LABSPEAR = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    strRegDate = Mid(strBarcode, 1, 8)
    strRegno = Mid(strBarcode, 9)
    
    If strBarcode = "" Then
        Exit Function
    End If
              
    If strRegno = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
        
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       CONVERT(NVARCHAR(50),M.접수일자,112)    AS HOSPDATE " & vbCrLf
    SQL = SQL & "     , M.접수번호                              AS PID      " & vbCrLf
    SQL = SQL & "     , M.차트번호                              AS CHARTNO  " & vbCrLf
    SQL = SQL & "     , M.성명                                  AS PNAME    " & vbCrLf
    SQL = SQL & "     , M.성별                                  AS SEX      " & vbCrLf
    SQL = SQL & "     , M.나이                                  AS AGE      " & vbCrLf
    SQL = SQL & "     , M.거래처명                              AS DEPT     " & vbCrLf
    SQL = SQL & "     , E.검사코드                              AS ITEM     " & vbCrLf
    SQL = SQL & "  FROM VW_검사접수 M, VW_검사결과 R, VW_검사코드 E         " & vbCrLf
    SQL = SQL & " WHERE M.접수일자      = CONVERT(DATE, '" & strRegDate & "')" & vbCrLf
    SQL = SQL & "   AND M.접수번호      = '" & strRegno & "'                " & vbCrLf
    SQL = SQL & "   AND E.학부코드      = '" & gHOSP.PARTCD & "'            " & vbCrLf    'U2
    SQL = SQL & "   AND E.검사코드      IN (" & gAllTestCd & ")             " & vbCrLf
    SQL = SQL & "   AND ISNULL(R.보고여부, 'N') <> 'Y'                      " & vbCrLf
    SQL = SQL & "   AND (R.결과값 IS NULL OR R.결과값 = '')                 " & vbCrLf
    SQL = SQL & "   AND M.접수일자      = R.접수일자                        " & vbCrLf
    SQL = SQL & "   AND M.접수번호      = R.접수번호                        " & vbCrLf
    SQL = SQL & "   AND R.검사코드      = E.검사코드                        " & vbCrLf
   
    Call SetSQLData("바코드조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    
    SetText SPD, "0", asRow, colCHECKBOX
        
    ReDim Preserve gPatTest(RS.RecordCount)
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
'                SetText SPD, strBarcode, asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("CHARTNO")) & "", asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                SetText SPD, Trim(RS.Fields("DEPT")) & "", asRow, colDEPT
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = strBarcode 'Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '환자 성별/나이
                With mPatient
                    .SEX = Trim(RS.Fields("SEX")) & ""
                    .AGE = Trim(RS.Fields("AGE")) & ""
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        
                        '-- 처방코드
                        'gArrEQP(intCol - colSTATE, 16) = Trim(RS.Fields("ORDCODE")) & ""
                        
                        Exit For
                    End If
                Next
                
                gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("ITEM")) & "',"
            End With
            DoEvents
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_LABSPEAR = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_LABSPEAR = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetSampleInfo_LABSPEAR" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_BIT70(ByVal asRow As Long, ByVal SPD As Object) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo            As Long
    
On Error GoTo DBErr
    
    GetSampleInfo_BIT70 = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
        
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       L.LABODRDTE     AS      HOSPDATE            " & vbCrLf
    SQL = SQL & "     , L.LABBARCOD     AS      BARCODE             " & vbCrLf
    SQL = SQL & "     , L.LABCHTNUM     AS      CHARTNO             " & vbCrLf
    SQL = SQL & "     , L.LABATTEND     AS      PID                 " & vbCrLf
    SQL = SQL & "     , M.MANADMFOR     AS      INOUT               " & vbCrLf
    SQL = SQL & "     , M.MANRESNUM     AS      JUMIN               " & vbCrLf
    SQL = SQL & "     , M.MANPATNAM     AS      PNAME               " & vbCrLf
    SQL = SQL & "     , L.LABODRSTP     AS      ORDCODE             " & vbCrLf
    SQL = SQL & "     , L.LABODRCOD     AS      ITEM                " & vbCrLf
    SQL = SQL & "  FROM ME_LABDAT   L                               " & vbCrLf
    SQL = SQL & "     , ME_DAT      D                               " & vbCrLf
    SQL = SQL & "     , ME_MAN      M                               " & vbCrLf
    SQL = SQL & " WHERE L.LABODRDTE = '" & strRegDate & "'          " & vbCrLf
    SQL = SQL & "   AND L.LABCHTNUM = '" & strChartNo & "'          " & vbCrLf
    SQL = SQL & "   AND L.LABKEYNUM = D.DATKEYNUM                   " & vbCrLf      '-- 테이블연결키값
    SQL = SQL & "   AND L.LABATTEND = D.DATATTEND                   " & vbCrLf      '-- 내원번호
    SQL = SQL & "   AND L.LABATTEND = M.MANATTEND                   " & vbCrLf      '-- 내원번호
    SQL = SQL & "   AND L.LABCHTNUM = D.DATCHTNUM                   " & vbCrLf      '-- 챠트번호
    SQL = SQL & "   AND L.LABCHTNUM = M.MANCHTNUM                   " & vbCrLf      '-- 챠트번호
    SQL = SQL & "   AND L.LABODRDTE = D.DATODRDTE                   " & vbCrLf      '-- 처방일자
    SQL = SQL & "   AND L.LABODRCOD IN (" & gAllTestCd & ")         " & vbCrLf
    SQL = SQL & "   AND (L.LABCANCEL = '' OR L.LABCANCEL IS NULL)   " & vbCrLf      '-- 취소여부
    SQL = SQL & "   AND (L.LABRESULT = '' OR L.LABRESULT IS NULL)   " & vbCrLf
    SQL = SQL & "   AND L.LABENDDEP < '3'                           " & vbCrLf      '-- 처리상태 (2:접수, 3:결과입력)
        
        
    Call SetSQLData("바코드조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    
    SetText SPD, "0", asRow, colCHECKBOX
        
    ReDim Preserve gPatTest(RS.RecordCount)
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("CHARTNO")) & "", asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("BARCODE")) & "", asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("JUMIN")) & "", asRow, colPAGE
                Select Case Trim(Trim(RS.Fields("INOUT")) & "")
                    Case "A":   SetText SPD, "외래", asRow, colINOUT
                    Case "F":   SetText SPD, "입원", asRow, colINOUT
                    Case Else:  SetText SPD, "", asRow, colINOUT
                End Select
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '환자 성별/나이
                With mPatient
                    '-- 테스트 후 적용
                    'Call CalAgeSex(Trim(RS.Fields("JUMIN")) & "", Format(Now, "yyyy/mm/dd"))
                    
                    '.SEX = Trim(RS.Fields("SEX")) & ""
                    '.AGE = Trim(RS.Fields("AGE")) & ""
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        
                        '-- 처방코드
                        gArrEQP(intCol - colSTATE, 16) = Trim(RS.Fields("ORDCODE")) & ""
                        
                        Exit For
                    End If
                Next
                
                gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("ITEM")) & "',"
                'gPatTest(intTestCnt) = Trim(RS.Fields("ITEM"))
            End With
            DoEvents
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_BIT70 = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_BIT70 = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetSampleInfo_BIT70" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_BIT(ByVal asRow As Long, ByVal SPD As Object) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo            As Long
    
On Error GoTo DBErr
    
    GetSampleInfo_BIT = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
        
'    SQL = ""
'    SQL = SQL & "SELECT DISTINCT "
'    SQL = SQL & "       SUBSTRING(R.RESACPDTM,1,8)  AS HOSPDATE     " & vbCrLf
'    SQL = SQL & "     , R.RESSPMNUM                 AS BARCODE      " & vbCrLf 'RESOCMNUM
'    SQL = SQL & "     , R.RESCHTNUM                 AS CHARTNO      " & vbCrLf
'    SQL = SQL & "     , P.PBSPATNAM                 AS PNAME        " & vbCrLf
'    SQL = SQL & "     , R.RESLABCOD                 AS ITEM         " & vbCrLf
'    SQL = SQL & "  FROM RESINF AS R                                 " & vbCrLf
'    SQL = SQL & "     , PBSINF AS P                                 " & vbCrLf
'    SQL = SQL & "WHERE R.RESSPMNUM  = '" & strBarcode & "'          " & vbCrLf
'    SQL = SQL & "  AND R.RESCHTNUM  = P.PBSCHTNUM                   " & vbCrLf
'    SQL = SQL & "  AND R.RESLABCOD IN (" & gAllTestCd & ")          " & vbCrLf
'    SQL = SQL & "  AND (R.RESREPTYP IS NULL OR R.RESREPTYP <> 'F')  " & vbCrLf         '--  'I':중간 'F' 완료"
'    SQL = SQL & "  AND (R.RESRLTVAL = ''  OR R.RESRLTVAL IS NULL)   " & vbCrLf
    
    SQL = ""
    SQL = SQL & " Select                                                    "
    SQL = SQL & "      P.PbsPatNam AS PNAME                                 " & vbCrLf
    SQL = SQL & "    , P.PbsResNum AS PJUMIN                                " & vbCrLf
    SQL = SQL & "    , W.OcmAcpDtm AS HOSPDATE                              " & vbCrLf
    SQL = SQL & "    , R.ResOcmNum AS BARCODE                               " & vbCrLf
    SQL = SQL & "    , W.OcmChtNum AS CHARTNO                               " & vbCrLf
    SQL = SQL & "    , R.ResLabCod AS ITEM                                  " & vbCrLf
    SQL = SQL & "    From OcmInf AS W                                       " & vbCrLf
    SQL = SQL & "       , OdrInf AS O                                       " & vbCrLf
    SQL = SQL & "       , ResInf AS R                                       " & vbCrLf
    SQL = SQL & "       , PbsInf AS P                                       " & vbCrLf
    SQL = SQL & "       , LabMst AS E                                       " & vbCrLf
    SQL = SQL & "    where W.OcmNum    = R.ResOcmNum                        " & vbCrLf
    SQL = SQL & "      and W.OcmChtNum = P.PbsChtNum                        " & vbCrLf
    SQL = SQL & "      and W.OcmNum    = O.OdrOcmNum                        " & vbCrLf
    SQL = SQL & "      And O.Odrdelflg = 'N'                                " & vbCrLf
    SQL = SQL & "      and O.OdrSeq = R.ResOdrSeq                           " & vbCrLf
    SQL = SQL & "      and R.ResLabCod = E.LabCod                           " & vbCrLf
    SQL = SQL & "      and Upper(R.ResLabCod) IN (" & UCase(gAllTestCd) & ")" & vbCrLf
    SQL = SQL & "      and W.OcmComStt Not In ('CN', 'CR','VC')             " & vbCrLf
    SQL = SQL & "      and (R.ResRltVal is null or R.ResRltVal = '')        " & vbCrLf
    SQL = SQL & "      and R.ResOcmNum = '" & strBarcode & "'               " & vbCrLf
    
    Call SetSQLData("바코드조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    
    SetText SPD, "0", asRow, colCHECKBOX
        
    ReDim Preserve gPatTest(RS.RecordCount)
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("CHARTNO")) & "", asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("BARCODE")) & "", asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("CHARTNO")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        
                        '-- 처방코드
                        'gArrEQP(intCol - colSTATE, 16) = Trim(RS.Fields("ORDCODE")) & ""
                        
                        Exit For
                    End If
                Next
                
                gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("ITEM")) & "',"
            End With
            DoEvents
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_BIT = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_BIT = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetSampleInfo_BIT" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_MCC(ByVal asRow As Long, ByVal SPD As Object) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo            As Long
    
On Error GoTo DBErr
    
    GetSampleInfo_MCC = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    If IsDate(strRegDate) Then
        strRegDate = Format(strRegDate, "yyyymmdd")
    End If
    
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    If Not IsNumeric(strBarcode) Then
        Exit Function
    End If
    
    
    
    Screen.MousePointer = 11
        
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       READING_YMD     AS HOSPDATE         " & vbCrLf
    SQL = SQL & "     , BCODE_NO        AS BARCODE          " & vbCrLf
    SQL = SQL & "     , PTNT_NO         AS PID              " & vbCrLf
    SQL = SQL & "     , PTNT_NM         AS PNAME            " & vbCrLf
    SQL = SQL & "     , AGE             AS AGE              " & vbCrLf
    SQL = SQL & "     , SEX             AS SEX              " & vbCrLf
    SQL = SQL & "     , IO_GB           AS INOUT            " & vbCrLf
    SQL = SQL & "     , ORD_CD          AS ITEM             " & vbCrLf
    SQL = SQL & "     , SP_CD           AS SPCCD            " & vbCrLf
    SQL = SQL & "  FROM LIS_INTERFACE1_V                    " & vbCrLf
    SQL = SQL & " WHERE BCODE_NO    = '" & strBarcode & "'  " & vbCrLf
    If strRegDate <> "" And IsDate(strRegDate) Then
        SQL = SQL & "   AND READING_YMD = '" & strRegDate & "'  " & vbCrLf
    End If
'    SQL = SQL & "   AND STS_CD = '0'                        " & vbCrLf  '0 접수, 1:결과전송
    If gAllTestCd <> "" Then
        SQL = SQL & "   AND ORD_CD IN (" & gAllTestCd & ")      " & vbCrLf
    End If
    
    SQL = SQL & " ORDER BY ORD_CD                           " & vbCrLf
        
    Call SetSQLData("바코드조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    
    SetText SPD, "0", asRow, colCHECKBOX
        
    'ReDim Preserve gPatTest(RS.RecordCount)
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Format(Trim(RS.Fields("HOSPDATE")) & "", "####-##-##"), asRow, colHOSPDATE
                SetText SPD, IIf(Trim(RS.Fields("INOUT")) & "" = "10", "입원", "외래"), asRow, colINOUT
                SetText SPD, Trim(RS.Fields("BARCODE")), asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '환자 성별/나이
                With mPatient
                    .AGE = Trim(RS.Fields("AGE")) & ""
                    .SEX = Trim(RS.Fields("SEX")) & ""
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        Exit For
                    End If
                Next
                
                gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("ITEM")) & "',"
                'gPatTest(intTestCnt) = Trim(RS.Fields("ITEM"))
            End With
            DoEvents
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_MCC = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_MCC = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "GetSampleInfo_MCC" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_NEOSENSE(ByVal asRow As Long, ByVal SPD As Object) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo        As Long
    
    Dim strYY           As String
    Dim strMM           As String
    Dim strDD           As String

On Error GoTo DBErr
    
    GetSampleInfo_NEOSENSE = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strChartNo = "" Then
        Exit Function
    End If
    
    If strRegDate = "" Then
        Exit Function
    End If
    
    strYY = Mid(strRegDate, 1, 4)
    strMM = Mid(strRegDate, 5, 2)
    strDD = Mid(strRegDate, 7, 2)
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       a.챠트번호          AS  CHARTNO " & vbCrLf
    SQL = SQL & "     , a.진료년            AS  HOSPYY  " & vbCrLf
    SQL = SQL & "     , a.진료월            AS  HOSPMM  " & vbCrLf
    SQL = SQL & "     , a.진료일            AS  HOSPDD  " & vbCrLf
    SQL = SQL & "     , b.수진자명          AS  PNAME   " & vbCrLf
    SQL = SQL & "     , b.주민등록번호      AS  JUMIN   " & vbCrLf
    SQL = SQL & "     , a.처방코드 + '|' + a.서브코드 AS ITEM   " & vbCrLf
    SQL = SQL & "  FROM TB_검사항목 a                   " & vbCrLf
    SQL = SQL & "     , TB_인적사항 b                   " & vbCrLf
    SQL = SQL & "     , tb_진료기본 c                   " & vbCrLf
    SQL = SQL & "     , tb_진료내역 d                   " & vbCrLf
    SQL = SQL & " WHERE a.챠트번호 = '" & strChartNo & "'   " & vbCrLf
    SQL = SQL & "   And a.진료년 = '" & strYY & "'    " & vbCrLf
    SQL = SQL & "   And a.진료월 = '" & strMM & "'    " & vbCrLf
    SQL = SQL & "   And a.진료일 = '" & strDD & "'    " & vbCrLf
    SQL = SQL & "   And a.처방번호 > 0                                          " & vbCrLf
    SQL = SQL & "   And c.진료상태 in( '1','5','6','7','8','9')                 " & vbCrLf
    SQL = SQL & "   And a.처방코드 + '|' + a.서브코드 IN (" & gAllTestCd & ")         " & vbCrLf
    SQL = SQL & "   And (a.검사결과 is null or a.검사결과 = '')                 " & vbCrLf
    SQL = SQL & "   And d.수정구분 <> '1'                                       " & vbCrLf '//진료내역삭제
    SQL = SQL & "   And a.진료년    = c.진료년          " & vbCrLf
    SQL = SQL & "   And a.진료월    = c.진료월          " & vbCrLf
    SQL = SQL & "   And a.진료일    = c.진료일          " & vbCrLf
    SQL = SQL & "   And a.챠트번호  = c.챠트번호        " & vbCrLf
    SQL = SQL & "   And a.챠트번호  = b.챠트번호        " & vbCrLf
    SQL = SQL & "   And a.진료년    = d.진료년          " & vbCrLf
    SQL = SQL & "   And a.진료월    = d.진료월          " & vbCrLf
    SQL = SQL & "   And a.진료일    = d.진료일          " & vbCrLf
    SQL = SQL & "   And a.챠트번호  = d.챠트번호        " & vbCrLf
    SQL = SQL & "   And a.처방코드  = d.처방코드        " & vbCrLf
    
    'SQL = SQL & "   And a.서브코드  = d.서브코드        " & vbCrLf
    'SQL = SQL & " Order By a.진료년, a.진료월, a.진료일, b.수진자명 " & vbCrLf
        
    Call SetSQLData("바코드조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    
    SetText SPD, "0", asRow, colCHECKBOX
        
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                SetText SPD, "1", asRow, colCHECKBOX
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        Exit For
                    End If
                Next
                
                gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("ITEM")) & "',"
            End With
            DoEvents
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_NEOSENSE = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_NEOSENSE = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetSampleInfo_NEOSENSE" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_SY(ByVal asRow As Long, ByVal SPD As Object) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo        As Long
    
    Dim strRegno        As String

On Error GoTo DBErr
    
    GetSampleInfo_SY = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    If Len(strBarcode) <> 13 Then
        Exit Function
    End If
    
    strRegDate = Mid(strBarcode, 1, 8)
    strRegno = Mid(strBarcode, 9)
    
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       CONVERT(NVARCHAR(50),M.접수일자,112)    AS HOSPDATE " & vbCrLf
    SQL = SQL & "     , M.접수번호                              AS PID      " & vbCrLf
    SQL = SQL & "     , M.차트번호                              AS CHARTNO  " & vbCrLf
    SQL = SQL & "     , M.성명                                  AS PNAME    " & vbCrLf
    SQL = SQL & "     , M.성별                                  AS SEX      " & vbCrLf
    SQL = SQL & "     , M.나이                                  AS AGE      " & vbCrLf
    SQL = SQL & "     , E.검사코드                              AS ITEM     " & vbCrLf
    SQL = SQL & "  FROM VW_검사접수 M, VW_검사결과 R, VW_검사코드 E         " & vbCrLf
    SQL = SQL & " WHERE M.접수일자 = CONVERT(DATE, '" & strRegDate & "')    " & vbCrLf
    SQL = SQL & "   AND M.접수번호 = '" & strRegno & "'                     " & vbCrLf
    SQL = SQL & "   AND E.학부코드 = '" & gHOSP.PARTCD & "'                 " & vbCrLf
    SQL = SQL & "   AND E.검사코드 IN (" & gAllTestCd & ")                  " & vbCrLf
    SQL = SQL & "   AND isnull(R.보고여부, 'N') <> 'Y'                      " & vbCrLf
    SQL = SQL & "   AND (R.결과값 is null or R.결과값 = '')                 " & vbCrLf
    SQL = SQL & "   AND M.접수일자 = R.접수일자                             " & vbCrLf
    SQL = SQL & "   AND M.접수번호 = R.접수번호                             " & vbCrLf
    SQL = SQL & " ORDER BY HOSPDATE, PID                                    " & vbCrLf
    
    Call SetSQLData("바코드조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    
    SetText SPD, "0", asRow, colCHECKBOX
        
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("CHARTNO")) & "", asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '환자 성별/나이
                With mPatient
                    .AGE = Trim(RS.Fields("AGE")) & ""
                    .SEX = Trim(RS.Fields("SEX")) & ""
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) & "" = gArrEQP(intCol - colSTATE, 6) Then
                        '.Row = asRow
                        '.Col = intCol
                        '.BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        Exit For
                    End If
                Next
                
                gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("ITEM")) & "',"
            End With
            DoEvents
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_SY = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_SY = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetSampleInfo_SY" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- 검사자 정보 가져오기
Function GetSampleInfo_EASYS(ByVal asRow As Long, ByVal SPD As Object) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo        As Long
    Dim strRegno        As String

    Dim strJuminSex     As String
    Dim srJumin         As String

On Error GoTo DBErr
    
    GetSampleInfo_EASYS = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    If IsDate(strRegDate) Then
        strRegDate = Format(strRegDate, "yyyymmdd")
    End If
    
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If Not IsNumeric(strBarcode) Then
        Exit Function
    End If
    
    If InStr(strBarcode, "+") > 0 Then
        Exit Function
    End If
    
    If InStr(strBarcode, "-") > 0 Then
        Exit Function
    End If
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       a.ACC_YMD   AS HOSPDATE " & vbCrLf
    SQL = SQL & "     , a.RECEPT_NO AS BARCODE  " & vbCrLf
    SQL = SQL & "     , a.PTNT_NO   AS PID      " & vbCrLf
    SQL = SQL & "     , c.PTNT_NM   AS PNAME    " & vbCrLf
    SQL = SQL & "     , c.BIRTH_YMD AS AGE      " & vbCrLf
    SQL = SQL & "     , c.SEX       AS SEX      " & vbCrLf
    SQL = SQL & "     , a.ORD_CD    AS ITEM     " & vbCrLf
    SQL = SQL & "  FROM H3LAB_RESULT a          " & vbCrLf
    SQL = SQL & "     , H1OPDIN b               " & vbCrLf
    SQL = SQL & "     , HZ_MST_PTNT c           " & vbCrLf
    SQL = SQL & " WHERE a.RECEPT_NO = '" & strBarcode & "'  " & vbCrLf
    If gAllTestCd <> "" Then
        SQL = SQL & "   AND a.ORD_CD IN (" & gAllTestCd & ")" & vbCrLf
    End If
'    SQL = SQL & "   AND a.STS_CD    = 'A'                   " & vbCrLf    'A:접수, R:결과전송
'    SQL = SQL & "   AND a.SUTAK_CD  = ''                    " & vbCrLf
    SQL = SQL & "   AND (a.RESULT_NM = '' OR a.RESULT_NM IS NULL)            " & vbCrLf
    SQL = SQL & "   AND a.RECEPT_NO = b.RECEPT_NO           " & vbCrLf
    SQL = SQL & "   AND a.PTNT_NO    = c.PTNT_NO            " & vbCrLf
    SQL = SQL & " ORDER BY a.ORD_CD                         " & vbCrLf
        
    Call SetSQLData("바코드조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    
    SetText SPD, "0", asRow, colCHECKBOX
        
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("BARCODE")), asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '환자 성별/나이
'                With mPatient
'                    .AGE = Trim(RS.Fields("AGE")) & ""
'                    .SEX = Trim(RS.Fields("SEX")) & ""
'                End With
                
    
                Select Case Trim(RS.Fields("SEX")) & ""
                Case "M"
                            If Mid(Trim(RS.Fields("AGE")) & "", 1, 1) = "1" Then
                                strJuminSex = "1"
                            ElseIf Mid(Trim(RS.Fields("AGE")) & "", 1, 1) = "2" Then
                                strJuminSex = "3"
                            End If
                Case "F"
                            If Mid(Trim(RS.Fields("AGE")) & "", 1, 1) = "1" Then
                                strJuminSex = "2"
                            ElseIf Mid(Trim(RS.Fields("AGE")) & "", 1, 1) = "2" Then
                                strJuminSex = "4"
                            End If
                End Select
                
                srJumin = Mid(Trim(RS.Fields("AGE")) & "", 3, 6) & strJuminSex & "000000"
                
                Call CalAgeSex(srJumin, Format(Now, "yyyy/mm/dd"))
                
                SetText SPD, mPatient.SEX, asRow, colPSEX
                SetText SPD, mPatient.AGE, asRow, colPAGE
                
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If GetTestNm(Trim(RS.Fields("ITEM")) & "", False) = gArrEQP(intCol - colSTATE, 6) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        Exit For
                    End If
                Next
                
                gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("ITEM")) & "',"
            End With
            DoEvents
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_EASYS = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_EASYS = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetSampleInfo_EASYS" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function


'-- 검사자 정보 가져오기
Function GetSampleInfo_UBCARE(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    
    
On Error GoTo DBErr
    
    GetSampleInfo_UBCARE = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
        
    SQL = ""
    SQL = SQL & "Select DISTINCT "
    SQL = SQL & "       SAVESEQ                 " & vbCrLf
    SQL = SQL & "     , HOSPDATE                " & vbCrLf
    SQL = SQL & "     , INOUT                   " & vbCrLf
    SQL = SQL & "     , CHARTNO                 " & vbCrLf
    SQL = SQL & "     , BARCODE                 " & vbCrLf
    SQL = SQL & "     , PID                     " & vbCrLf
    SQL = SQL & "     , PNAME                   " & vbCrLf
    SQL = SQL & "     , PSEX                    " & vbCrLf
    SQL = SQL & "     , PAGE                    " & vbCrLf
    SQL = SQL & "     , PJUMIN                  " & vbCrLf
    SQL = SQL & "     , EXAMCODE        AS ITEM " & vbCrLf
    SQL = SQL & "  From UB_PATRESULT                        " & vbCrLf
    SQL = SQL & " Where BARCODE = '" & strBarcode & "'      " & vbCrLf
    If strPatID <> "" Then
        SQL = SQL & "   And PID     = '" & strPatID & "'    " & vbCrLf
    End If
    SQL = SQL & "   And EXAMCODE IN (" & gAllTestCd & ")    " & vbCrLf
    If frmInterface.chkSave.Value = "0" Then
        SQL = SQL & "   And (RESULT = '' OR RESULT IS NULL) " & vbCrLf
        SQL = SQL & "   AND SENDFLAG = '0'                  " & vbCrLf
    End If
    SQL = SQL & " Order by SAVESEQ,HOSPDATE,PNAME           " & vbCrLf
        
    Call SetSQLData("바코드조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn_Local.CursorLocation = adUseClient
    Set RS = AdoCn_Local.Execute(SQL, , 1)
    
    SetText SPD, "0", asRow, colCHECKBOX
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                SetText SPD, "1", asRow, colCHECKBOX
                SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", asRow, colHOSPDATE
                Select Case Trim(Trim(RS.Fields("INOUT")) & "")
                    Case "0":   SetText SPD, "외래", asRow, colINOUT
                    Case "1":   SetText SPD, "입원", asRow, colINOUT
                    Case Else:  SetText SPD, Trim(Trim(RS.Fields("INOUT")) & ""), asRow, colINOUT
                End Select
                SetText SPD, Trim(RS.Fields("BARCODE")), asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("CHARTNO")), asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("PJUMIN")) & "", asRow, colPJUMIN
                SetText SPD, Trim(RS.Fields("PSEX")) & "", asRow, colPSEX
                SetText SPD, Trim(RS.Fields("PAGE")) & "", asRow, colPAGE
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                 
                mPatient.AGE = Trim(RS.Fields("PAGE")) & ""
                mPatient.SEX = Trim(RS.Fields("PSEX")) & ""
                
                '오더정보에 저장
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Seq = asRow
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) & "" = gArrEQP(intCol - colSTATE, 6) Then
                        '.Row = asRow
                        '.Col = intCol
                        '.BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        Exit For
                    End If
                Next
                
                gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("ITEM")) & "',"
                
            End With
            DoEvents
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_UBCARE = 1
    
    Screen.MousePointer = 0
    
    
Exit Function

DBErr:
    GetSampleInfo_UBCARE = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetSampleInfo_UBCARE" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function


'-- 검사자 정보 가져오기
Function GetSampleInfo_BITJSON(ByVal asRow As Long, ByVal SPD As Object) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo        As Long
    
    Dim strRegno        As String
    Dim strParam(0)     As Variant
    Dim strReturn       As Variant
    Dim strJData        As Variant
    Dim i               As Integer
    
    Dim strDate     As String
    Dim strWN       As String
    Dim strPtNm     As String
    Dim strSex      As String
    Dim strAge      As String
    Dim strPID      As String
    Dim strDept     As String
    Dim strExmnCd   As String
    
    
On Error GoTo DBErr
    
    GetSampleInfo_BITJSON = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Call SetSQLData("바코드조회", strParam(0), "A")
    strParam(0) = strBarcode
    strReturn = JsonSend("PATLIST", strParam)
    Call SetSQLData("바코드조회", strReturn, "A")
    Call getJsonVar(CStr(strReturn))
        
    strJData = Split(mJsonData, vbCr)
        
        
    If mJsonData <> "" Then
        With SPD
            .ReDraw = False
            For i = 0 To UBound(strJData)
                Select Case Trim(mGetP(strJData(i), 1, "@"))
                    Case "rcpnDt":      strDate = Trim(mGetP(strJData(i), 2, "@"))
                    Case "workNo":      strWN = Trim(mGetP(strJData(i), 2, "@"))
                    Case "ptNm":        strPtNm = Trim(mGetP(strJData(i), 2, "@"))
                    Case "brcdLablNo":  strBarcode = Trim(mGetP(strJData(i), 2, "@"))
                    Case "sex":         strSex = Trim(mGetP(strJData(i), 2, "@"))
                    Case "age":         strAge = Trim(mGetP(strJData(i), 2, "@"))
                    Case "pid":         strPID = Trim(mGetP(strJData(i), 2, "@"))
                    Case "medDp":       strDept = Trim(mGetP(strJData(i), 2, "@"))
                    Case "exmnCd":      strExmnCd = Trim(mGetP(strJData(i), 2, "@"))
                    Case "spcmCd": '마지막
                            
                            SetText SPD, "1", asRow, colCHECKBOX
                            SetText SPD, Mid(strDate, 1, 10), asRow, colHOSPDATE
                            SetText SPD, strBarcode, asRow, colBARCODE
                            SetText SPD, strWN, asRow, colCHARTNO
                            SetText SPD, strPID, asRow, colPID
                            SetText SPD, strPtNm, asRow, colPNAME
                            SetText SPD, strSex, asRow, colPSEX
                            SetText SPD, strAge, asRow, colPAGE
                            
                            '-- 화면에 표시
                            For intCol = colSTATE + 1 To .MaxCols
                                'If Trim(strExmnCd) = gArrEQP(intCol - colSTATE, 2) Then
                                If GetTestNm(Trim(strExmnCd), False) = gArrEQP(intCol - colSTATE, 6) Then
                                    .Row = asRow
                                    .Col = intCol
                                    '.BackColor = vbYellow
                                    Call SetText(SPD, "◇", asRow, intCol)
                                    Exit For
                                End If
                            Next
                            gPatOrdCd = gPatOrdCd & "'" & Trim(strExmnCd) & "',"
                                    
                End Select
            Next
        End With
    End If
    
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_BITJSON = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_BITJSON = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetSampleInfo_BITJSON" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- 검사자 정보 가져오기
'    send = ""
'    send = send & "MSH|^~\&|HL7|MMS|||20180129212430||ORU^R01|1a082e2:10e59b48c04:-2cf9:27695009|P|2.3||||||8859/1" & vbCr
'    send = send & "PID|||201801290036^송혜경^861121^2^20180129^20180129^DefaultDomain^PI" & vbCr
'    send = send & "PV1||E|I0101" & vbCr
'    send = send & "OBR|1||||||20180129212430" & vbCr
'    send = send & "OBX|1|ST|WMD0910185||||||||R" & vbCr
'    send = send & "OBX|2|ST|WMD0910186||||||||R" & vbCr
'    send = send & "OBX|3|ST|WMD0910187||||||||R" & vbCr
'    strReturn = send
'Function GetSampleInfo_HEALTH(ByVal asRow As Long, ByVal SPD As Object) As Integer
'    Dim oSOAP       As MSSOAPLib30.SoapClient30
'    Dim strParam    As String
'    Dim strReturn   As Variant
'    Dim strHData    As Variant
'
'    Dim strRegDate      As String
'    Dim strBarcode      As String
'    Dim strPatID        As String
'    Dim strChartNo      As String
'    Dim intCol          As Integer
'    Dim intTestCnt      As Integer
'    Dim lngRegNo        As Long
'
'    Dim strRegno        As String
'    Dim i               As Integer
'    Dim strExmnCd   As String
'
'
'On Error GoTo DBErr
'
'    GetSampleInfo_HEALTH = -1
'
'    intTestCnt = 0
'    gPatOrdCd = ""
'
'    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
'
'    If strBarcode = "" Then
'        Exit Function
'    End If
'
'    If Len(strBarcode) = 10 Then
'        strBarcode = "20" & strBarcode
'    End If
'
'    Set oSOAP = New MSSOAPLib30.SoapClient30
'    oSOAP.ClientProperty("ServerHTTPRequest") = True
'    oSOAP.MSSoapInit gHEALTH.INITURL
'
'    strParam = ""
'    strParam = strParam & "MSH|^~\&|HL7|MMS|||1||ORU^R01|1a082e2:10e59b48c04:-2cf9:27695009|P|2.3||||||8859/1" & Chr(13) & Chr(10)
'    strParam = strParam & "PID|||" & strBarcode & "^" & gHOSP.MACHCD & "^" & gHOSP.USERID & "^DefaultDomain^PI" & Chr(13) & Chr(10)
'    strParam = strParam & "PV1||E|" & gHOSP.HOSPCD & Chr(13) & Chr(10)
'    strParam = strParam & "OBR|1||||||1" & Chr(13) & Chr(10)
'
'    strParam = Chr(11) & strParam & Chr(12) & Chr(13)
'
'    Call SetSQLData("바코드조회", strParam)
'    strParam = MakeB64(strParam)
'    strReturn = oSOAP.New_SelectOrder(strParam)
'    strReturn = MakeUB64(strReturn)
'    Call SetSQLData("바코드조회", strReturn, "A")
'    Set oSOAP = Nothing
'    DoEvents
'
'    SPD.MaxRows = 0
'
'    If strReturn = "" Then
'        '조회대상자 없음
'        frmInterface.lblComStatus.Caption = "워크리스트 조회 대상자가 없습니다."
'    Else
'        strHData = Split(strReturn, vbCr)
'
'        With SPD
'            .ReDraw = False
'            For i = 0 To UBound(strHData)
'                Select Case mGetP(strHData(i), 1, "|")
'                    Case "PID"
'                        SetText SPD, "1", asRow, colCHECKBOX
'                        SetText SPD, Trim(mGetP(mGetP(strHData(i), 4, "|"), 5, "^")), asRow, colHOSPDATE
'                        SetText SPD, Trim(mGetP(mGetP(strHData(i), 4, "|"), 1, "^")), asRow, colBARCODE
'                        SetText SPD, Trim(mGetP(mGetP(strHData(i), 4, "|"), 3, "^")), asRow, colCHARTNO
'                        SetText SPD, Trim(mGetP(mGetP(strHData(i), 4, "|"), 3, "^")), asRow, colPID
'                        SetText SPD, Trim(mGetP(mGetP(strHData(i), 4, "|"), 2, "^")), asRow, colPNAME
'
'                    Case "OBX"
'                        strExmnCd = Trim(mGetP(strHData(i), 4, "|"))
'                        '-- 화면에 표시
'                        For intCol = colSTATE + 1 To .MaxCols
'                            If Trim(strExmnCd) = gArrEQP(intCol - colSTATE, 2) Then
'                                .Row = asRow
'                                .Col = intCol
'                                '.BackColor = vbYellow
'                                Call SetText(SPD, "◇", asRow, intCol)
'                                Exit For
'                            End If
'                        Next
'                        gPatOrdCd = gPatOrdCd & "'" & Trim(strExmnCd) & "',"
'                End Select
'            Next
'        End With
'    End If
'
'    If gPatOrdCd <> "" Then
'        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
'    End If
'
'    GetSampleInfo_HEALTH = 1
'
'    Screen.MousePointer = 0
'
'Exit Function
'
'DBErr:
'    GetSampleInfo_HEALTH = -1
'    intTestCnt = 0
'    Screen.MousePointer = 0
'
'    strErrMsg = ""
'    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_GetSampleInfo_HEALTH" & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
'    frmErrMsg.txtErr = vbNewLine & strErrMsg
'    frmErrMsg.Show
'
'End Function


'-- 검사자 정보 가져오기
Function GetSampleInfo_TWIN(ByVal asRow As Long, ByVal SPD As Object) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo            As Long
    
On Error GoTo DBErr
    
    GetSampleInfo_TWIN = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    ReDim Preserve gPatTest(0)
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
        
'    SQL = ""
'    SQL = SQL & "SELECT DISTINCT "
'    SQL = SQL & "       READING_YMD     AS HOSPDATE     " & vbCrLf
'    SQL = SQL & "     , BCODE_NO        AS BARCODE      " & vbCrLf
'    SQL = SQL & "     , PTNT_NO         AS PID          " & vbCrLf
'    SQL = SQL & "     , PTNT_NM         AS PNAME        " & vbCrLf
'    SQL = SQL & "     , AGE             AS AGE          " & vbCrLf
'    SQL = SQL & "     , SEX             AS SEX          " & vbCrLf
'    SQL = SQL & "     , IO_GB           AS INOUT        " & vbCrLf
'    SQL = SQL & "     , ORD_CD          AS ITEM         " & vbCrLf
'    SQL = SQL & "     , SP_CD           AS SPCCD        " & vbCrLf
'    SQL = SQL & "  FROM LIS_INTERFACE1_V                " & vbCrLf
'    SQL = SQL & " WHERE BCODE_NO = '" & strBarcode & "' " & vbCrLf
'    SQL = SQL & "   AND STS_CD = '0'" & vbCrLf                      '0 접수, 1:결과전송
'    SQL = SQL & "   AND ORD_CD IN (" & gAllTestCd & ") " & vbCrLf
'    SQL = SQL & " ORDER BY ORD_CD " & vbCrLf
        
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       C.JOBDATE                               AS HOSPDATE     " & vbCrLf
    SQL = SQL & "     , C.SPECNO                                AS BARCODE      " & vbCrLf
    SQL = SQL & "     , C.PTNO                                  AS CHARTNO      " & vbCrLf
    SQL = SQL & "     , C.JOBNO                                 AS PID          " & vbCrLf
    SQL = SQL & "     , DECODE(C.GBIO,'I','입원','O','외래')    AS INOUT        " & vbCrLf
    SQL = SQL & "     , C.SNAME                                 AS PNAME        " & vbCrLf
    SQL = SQL & "     , C.SEX                                   AS SEX          " & vbCrLf
    SQL = SQL & "     , C.AGE                                   AS AGE          " & vbCrLf
    SQL = SQL & "     , A.MASTERCODE                            AS ORDERCODE    " & vbCrLf
    SQL = SQL & "     , A.SUBCODE                               AS ITEM         " & vbCrLf
    SQL = SQL & "  From TW_HSP_OCS.TWEXAM_RESULTC A                             " & vbCrLf
    SQL = SQL & "     , TW_HSP_OCS.TWEXAM_MASTER  B                             " & vbCrLf
    SQL = SQL & "     , TW_HSP_OCS.TWEXAM_SPECMST C                             " & vbCrLf
    SQL = SQL & " Where A.SPECNO = '" & strBarcode & "'                         " & vbCrLf
    'SQL = SQL & "   And B.EQUCODE1 = '" & gHOSP.MACHCD & "'                     " & vbCrLf '장비코드
    'SQL = SQL & "   AND A.MASTERCODE IN (" & gAllTestCd & ")                    " & vbCrLf
    SQL = SQL & "   AND A.SUBCODE IN (" & gAllTestCd & ")                       " & vbCrLf
    SQL = SQL & "   AND C.STATUS  <= '3'                                        " & vbCrLf '검사상태
    SQL = SQL & "   And C.SPECNO  = A.SPECNO                                    " & vbCrLf
    SQL = SQL & "   And A.MASTERCODE = B.MASTERCODE                             " & vbCrLf
    SQL = SQL & " ORDER BY C.JOBDATE, C.SPECNO                                  " & vbCrLf
        
        
    Call SetSQLData("바코드조회", SQL)
    
    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    
    SetText SPD, "0", asRow, colCHECKBOX
        
    ReDim Preserve gPatTest(RS.RecordCount)
    
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1
                SetText SPD, "1", asRow, colCHECKBOX
                'SetText SPD, Format(Trim(RS.Fields("HOSPDATE")) & "", "####-##-##"), asRow, colHOSPDATE
                SetText SPD, Mid(Trim(RS.Fields("HOSPDATE")) & "", 1, 10), asRow, colHOSPDATE
                SetText SPD, Trim(RS.Fields("INOUT")), asRow, colINOUT
                SetText SPD, Trim(RS.Fields("BARCODE")), asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("CHARTNO")) & "", asRow, colCHARTNO
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                
                '오더갯수
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '오더정보에 저장
                With mOrder
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '환자 성별/나이
                With mPatient
                    .AGE = Trim(RS.Fields("AGE")) & ""
                    .SEX = Trim(RS.Fields("SEX")) & ""
                End With
                
                '-- 화면에 표시
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "◇", asRow, intCol)
                        
                        '-- 처방코드
                        gArrEQP(intCol - colSTATE, 16) = Trim(RS.Fields("ORDERCODE")) & ""
                        
                        Exit For
                    End If
                Next
                
                gPatOrdCd = gPatOrdCd & "'" & Trim(RS.Fields("ITEM")) & "',"
                gPatTest(intTestCnt) = Trim(RS.Fields("ITEM"))
            End With
            DoEvents
            
            RS.MoveNext
        Loop
    End If
    
    RS.Close
            
    If gPatOrdCd <> "" Then
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    GetSampleInfo_TWIN = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_TWIN = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "GetSampleInfo_TWIN" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

Function SetJudge(asResult As String, asEquipCode As String)

    Select Case gEMR
        Case "AMIS"                         '아미스
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "EMEDI"                        '이메디
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "BIT"                          '비트
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)

        Case "BIT70"                        '비트 HIB70
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "EASYS"                        '이지스
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "EONM"                         '이온엠
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
            
        Case "GSEN"                         '지센커뮤니케이션즈(이챠트)
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "HWASAN"                       '화산
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "JAINCOM"                       '자인컴
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "JWINFO"                       '중외정보
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
            
        Case "KCHART"                       '다대소프트
                SetJudge = SetJudge_KCHART(asResult, asEquipCode)
        
        Case "KOMAIN"                       '중외정보
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
            
        Case "KYU"                          '건양대학교병원
                '워크리스트 기능없음
                'SetJudge =  SetJudge_KYU(asResult,asEquipCode)
        Case "MEDICHART"                    '메디챠트
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
            
        Case "MEDIIT"
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
            
        Case "MEDITOLISS"                    '
                SetJudge = SetJudge_MEDITOLISS(asResult, asEquipCode)
            
        Case "MSINFOTEC"                    'MS인포텍
                SetJudge = SetJudge_MSINFOTEC(asResult, asEquipCode)
                
    End Select
    
End Function

Function SetJudge_LOCAL(asResult As String, asEquipCode As String)
    Dim RS_L        As ADODB.Recordset
    Dim i As Integer
    Dim sLVal As String
    Dim sHVal As String
    Dim sEquipCode As String
    Dim sEquipRes As String
    Dim sResFlag As String
    
    
    sEquipRes = Trim(asResult)
    sEquipCode = Trim(asEquipCode)
    sResFlag = ""
    
    If sEquipCode = "" Then
        Exit Function
    End If
    
    If Not IsNumeric(sEquipRes) Then
        Exit Function
    End If
    
    SQL = ""
    SQL = SQL & "SELECT REFLOW, REFHIGH                     " & vbCr
    SQL = SQL & "  FROM EQPMASTER                           " & vbCr
    SQL = SQL & " WHERE EQUIPCD     = '" & gHOSP.MACHCD & "'" & vbCr
    SQL = SQL & "   AND RSLTCHANNEL = '" & sEquipCode & "'  " & vbCr

    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
        If IsNumeric(Trim(RS_L.Fields("REFLOW")) & "") = True And IsNumeric(Trim(RS_L.Fields("REFHIGH")) & "") = True Then
            sLVal = Trim(RS_L.Fields("REFLOW")) & ""
            sHVal = Trim(RS_L.Fields("REFHIGH")) & ""
            If CCur(sEquipRes) > CCur(sLVal) And CCur(sEquipRes) < CCur(sHVal) Then
                sResFlag = ""
            ElseIf CCur(sHVal) <= CCur(sEquipRes) Then
                sResFlag = "H"
            ElseIf CCur(sLVal) >= CCur(sEquipRes) Then
                sResFlag = "L"
            End If
        End If
    End If
 
    SetJudge_LOCAL = sResFlag
    
End Function

Function SetJudge_EASYS(asResult As String, asEquipCode As String) As String
    Dim RSJ         As ADODB.Recordset
    Dim strLow      As String
    Dim strHigh     As String
    
    SetJudge_EASYS = ""
    
          SQL = "Select REFLOW, REFHIGH                     " & vbCr
    SQL = SQL & "  From EQPMASTER                           " & vbCr
    SQL = SQL & " Where EQUIPCD  = '" & gHOSP.MACHCD & "'   " & vbCr
    SQL = SQL & "   And TESTCODE = '" & asEquipCode & "'    " & vbCr
    
    Set RSJ = New ADODB.Recordset
    Set RSJ = AdoCn_Local.Execute(SQL, , 1)
    If Not RSJ.EOF = True And Not RSJ.BOF = True Then
        strLow = Trim(RSJ.Fields("REFLOW") & "")
        strHigh = Trim(RSJ.Fields("REFHIGH") & "")
        
        If strLow <> "" And strHigh <> "" And asResult <> "" And IsNumeric(strLow) And IsNumeric(strHigh) And IsNumeric(asResult) Then
            If Val(asResult) > Val(strHigh) Then
                SetJudge_EASYS = "H"
            ElseIf Val(asResult) < Val(strLow) Then
                SetJudge_EASYS = "L"
            Else
                SetJudge_EASYS = " "
            End If
        Else
            SetJudge_EASYS = " "
        End If
    Else
        SetJudge_EASYS = ""
    End If
        
    RSJ.Close

End Function

Function SetJudge_MSINFOTEC(asResult As String, asEquipCode As String) As String
    Dim RSJ         As ADODB.Recordset
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    Dim strAge      As String
    Dim strSex      As String
    Dim strYY, strMM, strDD, strDate  As String
    
On Error GoTo ErrorTrap
    
    SetJudge_MSINFOTEC = ""
    
    asResult = Replace(asResult, "<", "")
    asResult = Replace(asResult, ">", "")
    
    strAge = mPatient.AGE
    strSex = mPatient.SEX
    
    If strAge <> "" Then
        If strAge <= 7 Then
            SQL = "Select YMAX as MAX, YMIN as MIN "
        Else
            If strSex = "M" Then
                     SQL = "Select MMAX as MAX, MMIN as MIN "
            Else
                     SQL = "Select WMAX as MAX, WMIN as MIN "
            End If
        End If
    Else
        SQL = "Select MMAX as MAX, MMIN as MIN "
    End If
    
    SQL = SQL & "  From emr.LABMAST"
    SQL = SQL & " Where ORCD =  '" & asEquipCode & "'"
    
    Set RSJ = AdoCn.Execute(SQL)
    Do Until RSJ.EOF
        If IsNumeric(asResult) And IsNumeric(RSJ.Fields("MAX")) And IsNumeric(RSJ.Fields("MIN")) Then
            If Val(asResult) > Val(RSJ.Fields("MAX")) Then
                SetJudge_MSINFOTEC = "H"
            ElseIf Val(asResult) < Val(RSJ.Fields("MIN")) Then
                SetJudge_MSINFOTEC = "L"
            Else
                SetJudge_MSINFOTEC = " "
            End If
        Else
            SetJudge_MSINFOTEC = " "
        End If
        RSJ.MoveNext
    
    Loop
    
    RSJ.Close

Exit Function

ErrorTrap:
    SetJudge_MSINFOTEC = ""
    
End Function

Function SetJudge_MEDITOLISS(asResult As String, asEquipCode As String) As String
    Dim RSJ         As ADODB.Recordset
    Dim strRefVal   As String
    
On Error GoTo ErrorTrap
    
    SetJudge_MEDITOLISS = ""
    
    SQL = ""
    SQL = SQL & "SELECT REFER_VALUE                                 " & vbCr
    SQL = SQL & "  FROM MEDITOLISS..TOTRES                          " & vbCr
    SQL = SQL & " WHERE REQUEST_DATE    = '" & mResult.RsltDate & "'" & vbCr
    SQL = SQL & "   AND EXAM_NO         = '" & mResult.BarNo & "'   " & vbCr
    SQL = SQL & "   AND EXAM_CODE       = '" & asEquipCode & "'     " & vbCr
    
    Set RSJ = AdoCn.Execute(SQL)
    Do Until RSJ.EOF
        strRefVal = RSJ.Fields("REFER_VALUE").Value & ""
        If IsNumeric(asResult) And Len(strRefVal) > 0 Then
            If Val(Trim$(asResult)) < Val(Mid(strRefVal, 1, InStr(strRefVal, "~") - 1)) Then
                SetJudge_MEDITOLISS = "L"
            ElseIf Val(Trim$(asResult)) > Val(Mid(strRefVal, InStr(strRefVal, "~") + 1)) Then
                SetJudge_MEDITOLISS = "H"
            Else
                SetJudge_MEDITOLISS = ""
            End If
        End If
    Loop
                
    RSJ.Close
    
Exit Function

ErrorTrap:
    SetJudge_MEDITOLISS = ""
    
End Function

Function SetJudge_KCHART(asResult As String, asEquipCode As String) As String
    Dim RS1         As ADODB.Recordset
    Dim sEquipCode  As String
    Dim sEquipRes   As String
    Dim sResFlag    As String
    Dim strRefL     As String
    Dim strRefH     As String
    
    
    sEquipRes = Trim(asResult)
    sEquipCode = Trim(asEquipCode)
    sResFlag = ""
    
    If sEquipCode = "" Then
        Exit Function
    End If
    
    strRefL = ""
    strRefH = ""
    
'    SQL = SQL & "  L.진료검사ID AS R, " & vbCrLf
'    SQL = SQL & "  L.진료지원ID AS P, " & vbCrLf

    '성인남 참고치0~참고치1,
    '성인여 참고치2~참고치3,
    '소아남 참고치4~참고치5,
    '소아여 참고치6~참고치7
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       A.환자성별 AS 성별                                          " & vbCr
    SQL = SQL & "     , L.참고치0, L.참고치1, L.참고치2, L.참고치3                  " & vbCr
    SQL = SQL & "     , L.참고치4, L.참고치5, L.참고치6, L.참고치7                  " & vbCr
    SQL = SQL & "     , (L.처방코드 + L.서브코드) AS ITEM                           " & vbCr
    SQL = SQL & "  FROM             TB_진료검사 L                                   " & vbCr
    SQL = SQL & "       INNER JOIN  TB_진료지원 J ON (L.진료지원ID = J.진료지원ID)  " & vbCr
    SQL = SQL & "       INNER JOIN  TB_진료일반 A ON (J.진료일자   = A.진료일자     " & vbCr
    SQL = SQL & "                                AND  J.챠트번호   = A.챠트번호     " & vbCr
    SQL = SQL & "                                AND  J.진료번호   = A.진료번호)    " & vbCr
    SQL = SQL & "  Where L.검체번호 = '" & mResult.BarNo & "'                       " & vbCr
    SQL = SQL & "    AND L.검사상태 < 5                                             " & vbCr
    SQL = SQL & "    AND (L.처방코드 + L.서브코드) = '" & sEquipCode & "'           " & vbCr
                                                                 

     Call SetSQLData("참고치조회", SQL)
     
     '-- Record Count 가져옴
     AdoCn.CursorLocation = adUseClient
     Set RS1 = AdoCn.Execute(SQL, , 1)
     If Not RS1.EOF = True And Not RS1.BOF = True Then
         Do Until RS1.EOF
            strRefL = ""
            strRefH = ""
            If Trim(RS1.Fields("성별")) & "" = "M" Then
                If Trim(RS1.Fields("참고치0")) & "" <> "" Then
                    strRefL = Trim(RS1.Fields("참고치0")) & ""
                    strRefH = Trim(RS1.Fields("참고치1")) & ""
                End If
            Else
                If Trim(RS1.Fields("성별")) & "" = "F" Then
                    If Trim(RS1.Fields("참고치2")) & "" <> "" Then
                        strRefL = Trim(RS1.Fields("참고치2")) & ""
                        strRefH = Trim(RS1.Fields("참고치3")) & ""
                    Else
                        strRefL = Trim(RS1.Fields("참고치0")) & ""
                        strRefH = Trim(RS1.Fields("참고치1")) & ""
                    End If
                End If
            End If
            RS1.MoveNext
        Loop
    
        If IsNumeric(sEquipRes) And IsNumeric(strRefL) = True And IsNumeric(strRefH) = True Then
            If CCur(sEquipRes) > CCur(strRefL) And CCur(sEquipRes) < CCur(strRefH) Then
                sResFlag = ""
            ElseIf CCur(strRefH) <= CCur(sEquipRes) Then
                sResFlag = "H"
            ElseIf CCur(strRefL) >= CCur(sEquipRes) Then
                sResFlag = "L"
            End If
        End If
    End If
    
    RS1.Clone
    
    SetJudge_KCHART = sResFlag
    
End Function


Function SetResult(asResult As String, asEquipCode As String)
    Dim RS_L        As ADODB.Recordset
    Dim i As Integer
    Dim sEquipCode As String
    Dim sEquipRes As String
    Dim sResult As String
    Dim sPoint As Integer
    Dim sResType As String
    
    
    sEquipRes = Trim(asResult)
    sEquipCode = Trim(asEquipCode)
    
    If sEquipCode = "" Then
        Exit Function
    End If
    
    SQL = ""
    SQL = SQL & "SELECT RESPREC, REFLOW, REFHIGH " & vbCr
    SQL = SQL & "  FROM EQPMASTER " & vbCr
    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "'" & vbCr
    SQL = SQL & "   AND RSLTCHANNEL = '" & sEquipCode & "'"

    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
        If IsNumeric(Trim(RS_L.Fields("RESPREC")) & "") = True Then
            sPoint = CInt(Trim(RS_L.Fields("RESPREC")))
            sResType = ""
            For i = 0 To sPoint
                If i = 0 Then
                    sResType = "#0"
                ElseIf i = 1 Then
                    sResType = sResType & ".0"
                Else
                    sResType = sResType & "0"
                End If
            Next
            sResult = Format(sEquipRes, sResType)
        Else
            sResult = sEquipRes
        End If
    End If
 
    SetResult = sResult
    
End Function

Function SetCutOffResult(asResult As String, asEquipCode As String, asTestCode As String) As String
    Dim RS_L        As ADODB.Recordset
    Dim i As Integer
    Dim sEquipCode As String
    Dim sEquipRes As String
    Dim sResult As String
'    Dim sPoint As Integer
'    Dim sResType As String
    
    Dim dblLow      As Double
    Dim dblHigh     As Double
    Dim strLComp    As String
    Dim strHComp    As String
    
    sResult = ""
    sEquipRes = Trim(asResult)
    sEquipCode = Trim(asEquipCode)
    
    If sEquipCode = "" Then
        Exit Function
    End If
    
    SQL = ""
    SQL = SQL & "SELECT RESULTTYPE, COLIN, COLCOMP, COLOUT, COHIN, COHCOMP, COHOUT, COMOUT   " & vbCr
    SQL = SQL & "  FROM EQPMASTER                                                " & vbCr
    SQL = SQL & " WHERE EQUIPCD     = '" & gHOSP.MACHCD & "'                     " & vbCr
    SQL = SQL & "   AND RSLTCHANNEL = '" & sEquipCode & "'                       " & vbCr
    SQL = SQL & "   AND TESTCODE    = '" & asTestCode & "'                       " & vbCr

    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
        If Trim(RS_L.Fields("COLCOMP") & "") <> "" And Trim(RS_L.Fields("COLIN") & "") <> "" Then
            If IsNumeric(Trim(RS_L.Fields("COLIN") & "")) Then
                dblLow = CCur(RS_L.Fields("COLIN"))
                strLComp = Trim(RS_L.Fields("COLCOMP"))
                If strLComp = "<" Then
                    If CCur(asResult) < dblLow Then
                        sResult = Trim(RS_L.Fields("COLOUT") & "")
                    Else
                        sResult = Trim(RS_L.Fields("COMOUT") & "")
                    End If
                ElseIf strLComp = "<=" Then
                    If CCur(asResult) <= dblLow Then
                        sResult = Trim(RS_L.Fields("COLOUT") & "")
                    Else
                        sResult = Trim(RS_L.Fields("COMOUT") & "")
                    End If
                End If
            End If
        ElseIf Trim(RS_L.Fields("COHCOMP") & "") <> "" And Trim(RS_L.Fields("COHIN") & "") <> "" Then
            If IsNumeric(Trim(RS_L.Fields("COHIN") & "")) Then
                dblHigh = CCur(RS_L.Fields("COHIN"))
                strHComp = Trim(RS_L.Fields("COHCOMP"))
                If strHComp = ">" Then
                    If CCur(asResult) < dblLow Then
                        sResult = Trim(RS_L.Fields("COHOUT") & "")
                    Else
                        sResult = Trim(RS_L.Fields("COMOUT") & "")
                    End If
                ElseIf strHComp = ">=" Then
                    If CCur(asResult) >= dblHigh Then
                        sResult = Trim(RS_L.Fields("COHOUT") & "")
                    Else
                        sResult = Trim(RS_L.Fields("COMOUT") & "")
                    End If
                End If
            End If
        End If
    End If
    
    If sResult <> "" Then
        Select Case Trim(RS_L.Fields("RESULTTYPE") & "")
            Case "변함없음"
                    sResult = Trim(asResult)
            Case "정량"
                    sResult = Trim(asResult)
            Case "정성"
                    sResult = Trim(sResult)
            Case "정량(정성)"
                    sResult = asResult & "(" & Trim(sResult) & ")"
            Case "정성(정량)"
                    sResult = sResult & "(" & Trim(asResult) & ")"
        End Select
    End If
    
    RS_L.Close
    
    SetCutOffResult = sResult
    
End Function


Function SetLocalDB(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    Dim strSaveSeq As String

    With frmInterface
        Select Case gEMR
            Case "UBCARE"
                sExamDate = Format(Now, "yyyymmdd")
                If Trim(GetText(.spdOrder, asRow1, colSAVESEQ)) = "" Then
                    Exit Function
                End If
                
                SQL = ""
                SQL = SQL & "UPDATE UB_PATRESULT SET " & vbCr
                SQL = SQL & "   SAVESEQ       = " & Trim(GetText(.spdOrder, asRow1, colSAVESEQ)) & vbCr
                SQL = SQL & " , EXAMDATE      = '" & sExamDate & "' " & vbCr
                SQL = SQL & " , DISKNO        = '" & Trim(GetText(.spdOrder, asRow1, colRACKNO)) & "'" & vbCr
                SQL = SQL & " , POSNO         = '" & Trim(GetText(.spdOrder, asRow1, colPOSNO)) & "'" & vbCr
                SQL = SQL & " , SEQNO         = " & Trim(GetText(.spdResult, asRow2, colRSEQNO)) & vbCr
                SQL = SQL & " , EQUIPRESULT   = '" & Trim(GetText(.spdResult, asRow2, colRMACHRESULT)) & "'" & vbCr
                SQL = SQL & " , RESULT        = '" & Trim(GetText(.spdResult, asRow2, colRLISRESULT)) & "'" & vbCr
                SQL = SQL & " WHERE EQUIPNO   = '" & gHOSP.MACHCD & "' " & vbCr
                SQL = SQL & "   AND HOSPDATE  = '" & Trim(GetText(.spdOrder, asRow1, colHOSPDATE)) & "' " & vbCr
                SQL = SQL & "   AND BARCODE   = '" & Trim(GetText(.spdOrder, asRow1, colBARCODE)) & "' " & vbCr
                SQL = SQL & "   AND EQUIPCODE = '" & Trim(GetText(.spdResult, asRow2, colRCHANNEL)) & "'" & vbCr
                SQL = SQL & "   AND EXAMCODE  = '" & Trim(GetText(.spdResult, asRow2, colRTESTCD)) & "'"
                
                If Not DBExec(AdoCn_Local, SQL) Then
                    Exit Function
                End If
            
            Case Else
                sExamDate = Format(Now, "yyyymmdd")
                If Trim(GetText(.spdOrder, asRow1, colSAVESEQ)) = "" Then
                    Exit Function
                End If
        
                SQL = ""
                SQL = SQL & "DELETE FROM PATRESULT " & vbCr
                SQL = SQL & " WHERE EQUIPNO     = '" & gHOSP.MACHCD & "' " & vbCrLf
                SQL = SQL & "   AND EXAMDATE    = '" & Trim(GetText(.spdOrder, asRow1, colEXAMDATE)) & "' " & vbCrLf
                'SQL = SQL & "   AND EXAMTIME    = '" & Trim(GetText(.spdOrder, asRow1, colEXAMTIME)) & "' " & vbCrLf
                'SQL = SQL & "   AND SAVESEQ     = " & Trim(GetText(.spdOrder, asRow1, colSAVESEQ)) & vbCrLf
                SQL = SQL & "   AND HOSPDATE    = '" & Trim(GetText(.spdOrder, asRow1, colHOSPDATE)) & "' " & vbCrLf
                SQL = SQL & "   AND BARCODE     = '" & Trim(GetText(.spdOrder, asRow1, colBARCODE)) & "' " & vbCrLf
                SQL = SQL & "   AND EQUIPCODE   = '" & Trim(GetText(.spdResult, asRow2, colRCHANNEL)) & "'" & vbCrLf
                SQL = SQL & "   AND EXAMCODE    = '" & Trim(GetText(.spdResult, asRow2, colRTESTCD)) & "'" & vbCrLf
        
                If DBExec(AdoCn_Local, SQL) Then
                    SQL = ""
                    SQL = SQL & "INSERT INTO PATRESULT (" & vbCrLf
                    SQL = SQL & "  EQUIPNO"                         '장비코드
                    SQL = SQL & ", EXAMDATE"                        '검사일자
                    SQL = SQL & ", EXAMTIME"                        '검사시간
                    SQL = SQL & ", SAVESEQ"                         '저장순번(날짜별)
                    SQL = SQL & ", HOSPDATE" & vbCrLf               '병원접수일자
                    
                    SQL = SQL & ", BARCODE"                         '검체번호
                    SQL = SQL & ", PID"                             '병록번호(내원번호)
                    SQL = SQL & ", CHARTNO"                         '챠트번호
                    SQL = SQL & ", SPECIMEN"                        '검체
                    SQL = SQL & ", DEPT" & vbCrLf                   '의뢰과
                    
                    SQL = SQL & ", INOUT"                           '입/외
                    SQL = SQL & ", ERYN"                            '응급여부
                    SQL = SQL & ", RETESTYN"                        '재검여부
                    SQL = SQL & ", PNAME"                           '이름
                    SQL = SQL & ", PSEX" & vbCrLf                   '성별(M,F)
                    
                    SQL = SQL & ", PAGE"                            '나이
                    SQL = SQL & ", EXAMUID"                         '검사자ID
                    SQL = SQL & ", DISKNO"                          'Rack
                    SQL = SQL & ", POSNO"                           'Pos
                    SQL = SQL & ", EQPSEQNO" & vbCrLf               '장비검사번호
                    '============================================================
                    
                    SQL = SQL & ", SEQNO"                           '검사일련번호
                    SQL = SQL & ", EQUIPCODE"                       '검사채널
                    SQL = SQL & ", ORDERCODE"                       '병원처방코드
                    SQL = SQL & ", EXAMCODE"                        '병원검사코드
                    SQL = SQL & ", EXAMCODESUB" & vbCrLf            '병원검사코드(SUB)"
                    
                    SQL = SQL & ", EXAMNAME"                        '검사명
                    SQL = SQL & ", EQUIPRESULT"                     '장비결과"
                    SQL = SQL & ", RESULT"                          '소수점적용"
                    SQL = SQL & ", PREVRESULT"                      '이전결과"
                    SQL = SQL & ", REFJUDGE" & vbCrLf               '판정(H,L)
                    
                    SQL = SQL & ", REFFLAG"                         'flag
                    SQL = SQL & ", REFVALUE"                        '참고치
                    SQL = SQL & ", PANICVALUE"                      'Delta
                    SQL = SQL & ", DELTAVALUE"                      'Panic
                    SQL = SQL & ", SENDFLAG"                        '전송구분(0:미전송,1:전송)"
                    SQL = SQL & ", SENDDATE)" & vbCrLf               '전송일자
                    
                    SQL = SQL & " VALUES (" & vbCrLf
                    SQL = SQL & "'" & gHOSP.MACHCD & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colEXAMDATE)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colEXAMTIME)) & "'"
                    SQL = SQL & "," & Trim(GetText(.spdOrder, asRow1, colSAVESEQ))
                    SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colHOSPDATE)) & "'" & vbCrLf
                    
                    SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colBARCODE)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colPID)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colCHARTNO)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colSPECIMEN)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colDEPT)) & "'" & vbCrLf
                    
                    SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colINOUT)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colER)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colRT)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colPNAME)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colPSEX)) & "'" & vbCrLf
                    
                    SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colPAGE)) & "'"
                    SQL = SQL & ",'" & gHOSP.USERID & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colRACKNO)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colPOSNO)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdOrder, asRow1, colSEQNO)) & "'" & vbCrLf
                    '============================================================
                    
                    SQL = SQL & "," & Trim(GetText(.spdResult, asRow2, colRSEQNO))
                    SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRCHANNEL)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRORDERCD)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRTESTCD)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRSUBCD)) & "'" & vbCrLf
                    
                    SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRTESTNM)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRMACHRESULT)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRLISRESULT)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRPREVRESULT)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRJUDGE)) & "'" & vbCrLf
                    
                    SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRFLAG)) & "'"
                    SQL = SQL & ",'" & Trim(GetText(.spdResult, asRow2, colRREF)) & "'"
                    SQL = SQL & ",''"
                    SQL = SQL & ",''"
                    SQL = SQL & ",'0'"
                    SQL = SQL & ",'')" & vbCrLf
                    
        '            Call SetSQLData("로컬저장", SQL)
                    
                    If Not DBExec(AdoCn_Local, SQL) Then
                        Exit Function
                    End If
        
                End If
        End Select
    End With

End Function

'-- 오늘 검사한 날짜의 Max + 1 번호를 가져온다
Public Function getMaxTestNum(ByVal strDate As String) As Long

    getMaxTestNum = 1

          SQL = "SELECT MAX(SAVESEQ) as SEQ FROM PATRESULT  "
    SQL = SQL & " WHERE EXAMDATE = '" & strDate & "' " & vbCrLf

    Set RS = AdoCn_Local.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        If Trim(RS.Fields("SEQ") & "") = "" Then
            getMaxTestNum = 1
        Else
            getMaxTestNum = Trim(RS.Fields("SEQ")) + 1
        End If
    Else
        getMaxTestNum = 1
    End If

    If getMaxTestNum >= 99999 Then
        getMaxTestNum = 99999
    End If

    RS.Close

End Function
'
