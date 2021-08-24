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

    '-- ������ �˻��ڵ��� ä�� ã��
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

    '-- ������ �˻��ڵ��� ä�� ã��
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

Public Function GetEquipExamCode_ATELLICA(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim strExamCode     As String
    Dim strSendCH       As String
    
    GetEquipExamCode_ATELLICA = ""
    strExamCode = ""
    mOrder.SendCnt = 0
    
    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If

    '-- ������ �˻��ڵ��� ä�� ã��
    SQL = ""
    SQL = SQL & "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"

    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    
    
    '^^^T3\^^^T4\^^^TSH
    
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strSendCH = Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
            If strSendCH <> "" Then
                mOrder.SendCnt = mOrder.SendCnt + 1
                If strExamCode = "" Then
                    strExamCode = "^^^" & strSendCH '& "^1"
                Else
                    strExamCode = strExamCode & "\^^^" & strSendCH '& "^1"
                End If
            End If
            
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close
    
    If strExamCode <> "" Then
        GetEquipExamCode_ATELLICA = strExamCode
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

    '-- ������ �˻��ڵ��� ä�� ã��
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

    '-- ������ �˻��ڵ��� ä�� ã��
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

    '-- ������ �˻��ڵ��� ä�� ã��
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

    '-- ������ �˻��ڵ��� ä�� ã��
    SQL = ""
    SQL = SQL & "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"

    AdoCn_Local.CursorLocation = adUseClient
    
    'CommandText        String  ������ ����� ����ϴ� �Ű������̸�, SQL ����, ���̺� ��, ���� ���ν����� ������ �� �ִ�.
    'RecordsAffected    Long    Execute �޼��忡 ���ؼ� ������ ���� ���ڵ��� ������ ��ȯ�Ѵ�. ���� ��� Delete������ �����ߴµ�, 10 ���� ���ڵ尡 �����Ǿ��ٸ�, 10 �̶�� ���� ��ȯ�Ѵ�.
    'Options            Long    Provider�� CommandText�� ��� ���������� �����ϴ� ����� �����ϴ� ���̸�, ������ ������ Long�̴�.
    '                    1      : adCmdText         CommandText�� ���� SQL �������� ó���Ѵ�.
    '                    2      : adCmdTable        CommandText�� ���� ���̺� ������ �ϴ� SQL ������ ���� ó���Ѵ�.
    '                    512    : adCmdTableDirect  CommandText�� ���� ���̺� ������ ó���Ѵ�.
    '                    4      : adCmdStoredProc   CommandText�� ���� ���� ���ν����� ó���Ѵ�.
    '                    8      : adCmdUnknown      ����� ������ �� �� �������� ó���Ѵ�.
    '                    16     : adAsyncExecute    ����� �񵿱������� �����Ѵ�.
    '                    32     : adAsyncFetch      CacheSize �Ӽ��� ������ �� ��ŭ�� ���ڵ徿 �񵿱������� ó���Ѵ�.
    
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


Public Function GetEquipExamCode_XN1000(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim strExamCode     As String
    Dim strSendCH       As String
    Dim strCBC      As String
    Dim strDIFF     As String
    
    GetEquipExamCode_XN1000 = ""
    strExamCode = ""

    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        '-- ������ ���� ��� CBC/ DIFF �˻��ϵ��� �Ѵ�.
        If strExamCode = "" Then
            strExamCode = "^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\^^^^RDW-SD\^^^^RDW-CV\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT\^^^^NRBC#\^^^^NRBC%\"
            strExamCode = strExamCode & "^^^^NEUT#\^^^^LYMPH%\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH#\^^^^LYMPH#\^^^^MONO%\^^^^EO%\^^^^BASO%\^^^^IG#\^^^^IG%\"
        End If
        
        If strExamCode <> "" Then
            strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
        End If
        
        GetEquipExamCode_XN1000 = strExamCode
        
        Exit Function
    End If

    '-- ������ �˻��ڵ��� ä�� ã��
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
                   
                    '-- ^^^^LYMPH#\�� �ΰ��� ������ ETB �� ��񿡼� �ν����� ���ϱ� ����..(�� �ڸ��� 230)
                    'strDIFF = "^^^^NEUT#\^^^^LYMPH%\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH#\^^^^LYMPH#\^^^^MONO%\^^^^EO%\^^^^BASO%\^^^^IG#\^^^^IG%\"
                    strDIFF = "^^^^NEUT#\^^^^LYMPH%\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH#\^^^^MONO%\^^^^EO%\^^^^BASO%\^^^^IG#\^^^^IG%\"
                    
                End If
            End If
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close
    
    strExamCode = strCBC & strDIFF
    
    '-- ������ ���� ��� CBC/ DIFF �˻��ϵ��� �Ѵ�.
    If strExamCode = "" Then
        strExamCode = "^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\^^^^RDW-SD\^^^^RDW-CV\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT\^^^^NRBC#\^^^^NRBC%\"
        strExamCode = strExamCode & "^^^^NEUT#\^^^^LYMPH%\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH#\^^^^LYMPH#\^^^^MONO%\^^^^EO%\^^^^BASO%\^^^^IG#\^^^^IG%\"
    End If
    
    If strExamCode <> "" Then
        GetEquipExamCode_XN1000 = Mid(strExamCode, 1, Len(strExamCode) - 1)
    End If
    
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
        '-- ������ ���� ��� CBC/DIFF �˻��ϵ��� �Ѵ�.
        If strExamCode = "" Then
            strExamCode = "^^^DIF"
        End If
        
        GetEquipExamCode_YUMIZEN = strExamCode
        
        Exit Function
    End If

    '-- ������ �˻��ڵ��� ä�� ã��
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
    
    '-- ������ ���� ��� CBC/ DIFF �˻��ϵ��� �Ѵ�.
    If strExamCode = "" Then
        strExamCode = "^^^DIF"
    End If
    
    If strExamCode <> "" Then
        GetEquipExamCode_YUMIZEN = strExamCode
    End If
    
End Function


'�� ���ä�ο� �˻��ڵ尡 1���̻� ���� (GLU-FBS, GLU-PP2..)
Public Function GetEquipExamCode_AU680(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim strExamCode     As String
    Dim strSendCH       As String
    
    GetEquipExamCode_AU680 = ""
    strExamCode = ""

    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If

    '-- ������ �˻��ڵ��� ä�� ã��
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

'�� ���ä�ο� �˻��ڵ尡 1���̻� ���� (GLU-FBS, GLU-PP2..)
Public Function GetEquipExamCode_HITACHI7180(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim strExamCode     As String
    Dim intIntBase      As Integer
    Dim strItems        As String           '������ �˻��׸�
    Dim blnISE          As Boolean          'Na, K, Cl �˻翩��


    strItems = String$(88, "0")
    
    GetEquipExamCode_HITACHI7180 = strItems
    strExamCode = ""
    blnISE = False
    mOrder.SendCnt = 0
    
    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If


    '-- ������ �˻��ڵ��� ä�� ã��
    SQL = ""
    SQL = SQL & "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"

    'Call SetSQLData("ä��", SQL)
    
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
'            intIntBase = AdoRs_Local.Fields("SENDCHANNEL").Value & ""
'            Mid$(strItems, intIntBase, 1) = "1"
'            mOrder.SendCnt = mOrder.SendCnt + 1


            If IsNumeric(AdoRs_Local.Fields("SENDCHANNEL").Value) Then

                intIntBase = CInt(AdoRs_Local.Fields("SENDCHANNEL").Value)
                'If intIntBase <> "" Then
                    '## ����׸�: 93~100
                    'If intIntBase >= 93 And intIntBase <= 100 Then
                    '    'GoTo Skip1
                    'Else
                        '## Na, K, Cl �˻翩�� Check
                        If intIntBase = 87 Or intIntBase = 88 Or intIntBase = 89 Then
                            blnISE = True
                        ElseIf intIntBase = 93 Then 'GA%
                            'GSP : 20
                            Mid$(strItems, 20, 1) = "1"
                            'GA-Alb : 21
                            Mid$(strItems, 21, 1) = "1"
                        Else
                            Mid$(strItems, intIntBase, 1) = "1"
                        End If
                        
                    'End If
                    mOrder.SendCnt = mOrder.SendCnt + 1
                'End If
            End If
            
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close

    'Call SetSQLData("ITEMS", strItems)
    
    GetEquipExamCode_HITACHI7180 = strItems
    

End Function

'�� ���ä�ο� �˻��ڵ尡 1���̻� ���� (GLU-FBS, GLU-PP2..)
Public Function GetEquipExamCode_HITACHI7020(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim strExamCode     As String
    Dim intIntBase      As Integer
    Dim strItems        As String           '������ �˻��׸�
    Dim blnISE          As Boolean          'Na, K, Cl �˻翩��

    strItems = String$(37, "0")
    
    GetEquipExamCode_HITACHI7020 = strItems
    strExamCode = ""
    blnISE = False
    mOrder.SendCnt = 0
    
    If Trim(argEquipCode) = "" Then
        Exit Function
    End If

    '-- ������ �˻��ڵ��� ä�� ã��
    SQL = ""
    SQL = SQL & "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"

    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            intIntBase = AdoRs_Local.Fields("SENDCHANNEL").Value & ""
            
            If intIntBase <> 99 Then
                Mid$(strItems, intIntBase, 1) = "1"
                mOrder.SendCnt = mOrder.SendCnt + 1
            End If
            AdoRs_Local.MoveNext
        Loop
    End If

'    '## Na, K, Cl �˻翩�� Check
'    If blnISE Then
'        Mid$(strItems, 87, 1) = "1"
'        mOrder.SendCnt = mOrder.SendCnt + 1
'    End If

    AdoRs_Local.Close

    GetEquipExamCode_HITACHI7020 = strItems
    
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

    With frmMain
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
        
        '-- Local���� ȯ�ں��� ����� ��������
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
                
                '-- ���������
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                  
                '-- �������� Ű ��������
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
                    '-- ��������
                    SQL = "" '
                    SQL = SQL & "Update TB_H141_LISTAKEBODY                     " & vbCrLf
                    SQL = SQL & "   SET H141_RSLTYN    ='Y'                     " & vbCrLf
                    SQL = SQL & " WHERE H141_TSAMPLENO = '" & strBarcode & "'   " & vbCrLf
                    SQL = SQL & "   AND H141_SUGACD    = '" & strTestCd & "'    " & vbCrLf
                    
                    Call SetSQLData("�������", SQL, "A")
                    AdoCn.Execute SQL
                    
                    SQL = ""
                    SQL = SQL & "UPDATE TB_H131_SPPRESULT                       " & vbCrLf
                    SQL = SQL & "   SET H131_RESULT  = '" & sResult & "'        " & vbCrLf
                    SQL = SQL & " WHERE H131_SPPTYPE = '" & gHOSP.PARTCD & "'   " & vbCrLf    'L010
                    SQL = SQL & "   AND H131_SEQNO   = '" & strTestCdSub & "'   " & vbCrLf
                        
                    Call SetSQLData("�������", SQL, "A")
                    AdoCn.Execute SQL
                
                    SQL = ""
                    SQL = SQL & "UPDATE TB_H130_SPPRECEIVE                              " & vbCrLf
                    SQL = SQL & "   SET H130_RSLTDAT = TO_CHAR(SYSDATE, 'YYYYMMDD')     " & vbCrLf
                    SQL = SQL & "      ,H130_RSLTTM  = TO_CHAR(SYSDATE, 'HH24:MI:SS')   " & vbCrLf
                    SQL = SQL & " WHERE H130_SPPTYPE = '" & gHOSP.PARTCD & "'           " & vbCrLf    'L010
                    SQL = SQL & "   AND H130_SEQNO   = '" & strTestCdSub & "'           " & vbCrLf
                        
                    Call SetSQLData("�������", SQL, "A")
                    AdoCn.Execute SQL
                
                    SQL = ""
                    SQL = SQL & "UPDATE TB_H140_LISTAKEHEAD                     " & vbCrLf
                    SQL = SQL & "   SET H140_RSLTYN    = 'Y'                    " & vbCrLf
                    SQL = SQL & " WHERE H140_TSAMPLENO = '" & strBarcode & "'   " & vbCrLf
                                        
                    Call SetSQLData("�������", SQL, "A")
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "SaveTransData_EONM" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
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

    With frmMain
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
        
        '-- Local���� ȯ�ں��� ����� ��������
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
                
                '-- ���������
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                  
                '-- �������� Ű ��������
                If strOrdCd = "" Then
                    strOrdCd = GetSampleSubITEM(strBarcode, strTestCd)
                End If
                
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" And strOrdCd <> "" Then
                    '-- ��������
                    SQL = ""
                    SQL = SQL & "Update RESULTOFNUM                                     " & vbCrLf
                    SQL = SQL & "   Set RESULTINDATE   = to_char(sysdate,'yyyymmdd')    " & vbCrLf
                    SQL = SQL & "     , RESULTINTIME   = to_char(sysdate,'HH24MI')      " & vbCrLf
                    SQL = SQL & "     , RESULTINID     = '" & gHOSP.USERID & "'         " & vbCrLf
                    SQL = SQL & "     , RESULTFLAG     = '1'                            " & vbCrLf
                    SQL = SQL & "     , TEXTRESULTVAL  = '" & sResult & "'              " & vbCrLf
                    '-- ����� ��ġ���̸�
                    If IsNumeric(sResult) Then
                        SQL = SQL & "     , NUMRESULTVAL = '" & sResult & "'           " & vbCrLf
                    End If
                    SQL = SQL & " Where SPCMNO         = '" & strBarcode & "'           " & vbCrLf
                    SQL = SQL & "   And ORDERCODE      = '" & strOrdCd & "'             " & vbCrLf
                    SQL = SQL & "   And RESULTITEMCODE = '" & strTestCd & "'            " & vbCrLf
                    SQL = SQL & "   And RESULTFLAG < '3'                                " & vbCrLf
                    
                    Call SetSQLData("�������", SQL, "A")
                    AdoCn.Execute SQL
                                        
                    '-- ���º���
                    SQL = ""
                    SQL = SQL & "Update REGISTINFOS                         " & vbCrLf
                    SQL = SQL & "   Set RESULTSTATE  = '1'                  " & vbCrLf
                    SQL = SQL & "      ,RsvAcptState = '4'                  " & vbCrLf
                    SQL = SQL & " Where SPCMNO       = '" & strBarcode & "' " & vbCrLf
                    SQL = SQL & "   AND ORDERCODE    = '" & strOrdCd & "'   " & vbCrLf
                    SQL = SQL & "   AND CLAS         = 4                    " & vbCrLf
                    SQL = SQL & "   AND RESULTSTATE < '4'                   " & vbCrLf
                    
                    Call SetSQLData("���º���", SQL, "A")
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SaveTransData_AMIS" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
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

    With frmMain
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
        
        '-- Local���� ȯ�ں��� ����� ��������
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
                
                '-- ���������
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                
                'MsgBox strOrdCd & "," & strTestCd & "," & strTestCdSub
                
                
                '-- �������� Ű ��������
                If strOrdCd = "" Then
                    strOrdCd = GetSampleSubITEM(strBarcode, strTestCd)
                    strOrdCd = mGetP(strOrdCd, 1, "|")
                    strTestCdSub = mGetP(strOrdCd, 2, "|")
                End If
                
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" And strOrdCd <> "" And strTestCdSub <> "" Then
                    '-- �������
                    'SQL = SQL & "    ,  ������ = 'IIS', " & vbCr
                    SQL = ""
                    SQL = SQL & "Update TB_����˻�                                   " & vbCrLf
                    SQL = SQL & "   Set �˻���              = '" & sResult & "'     " & vbCrLf
                    SQL = SQL & "     , ���̷ο�              = '" & strJudge & "'    " & vbCrLf
                    SQL = SQL & "     , �˻����              = '2'                   " & vbCrLf
                    SQL = SQL & "     , ��������              = '1'                   " & vbCrLf
                    SQL = SQL & "     , ��������              = GetDate()             " & vbCrLf
                    SQL = SQL & " Where ����˻�ID            = '" & strOrdCd & "'    " & vbCrLf
                    SQL = SQL & "   And ��������ID            = '" & strTestCdSub & "'" & vbCrLf
                    SQL = SQL & "   And ��ü��ȣ              = '" & strBarcode & "'  " & vbCrLf
                    SQL = SQL & "   And (ó���ڵ� + �����ڵ�) = '" & strTestCd & "'   " & vbCrLf
                    
                    Call SetSQLData("�������", SQL, "A")
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SaveTransData_KCHART" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
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

    With frmMain
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
        
        '-- Local���� ȯ�ں��� ����� ��������
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
                
                '-- ���������
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                
                '-- �������� Ű ��������
                If strOrdCd = "" Then
                    strOrdCd = GetSampleSubITEM(strBarcode, strTestCd)
                End If
                
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" And strOrdCd <> "" Then
                    '-- �������
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
                    If frmMain.chkSave.Value = "0" Then
                        SQL = SQL & "   And TRANSFLAG   < '2'                   " & vbCrLf
                    End If

                    
                    Call SetSQLData("�������", SQL, "A")
                    AdoCn.Execute SQL
                    
                    '-- ���º���
                          SQL = "Update SLA_LabMaster                       " & vbCrLf
                    SQL = SQL & "   Set JStatus = '2'                       " & vbCrLf
                    SQL = SQL & " Where SPECIMENNUM = '" & strBarcode & "'  " & vbCrLf
                    SQL = SQL & "   AND OrderCode   = '" & strOrdCd & "'    " & vbCrLf
                    If frmMain.chkSave.Value = "0" Then
                        SQL = SQL & "   And JStatus < '3'                       " & vbCrLf
                    End If
                    
                    Call SetSQLData("�������", SQL, "A")
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SaveTransData_JWINFO" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
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

    With frmMain
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
        
        '-- Local���� ȯ�ں��� ����� ��������
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
                
                '-- ���������
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                  
                '-- �������� Ű ��������
                If strOrdCd = "" Then
                    strOrdCd = GetSampleSubITEM(strBarcode, strTestCd)
                End If
                
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" And strOrdCd <> "" Then
                    '-- ��������
                    SQL = ""
                    SQL = SQL & "Update TB_�˻��׸� "
                    SQL = SQL & "   Set �˻���        = '" & sResult & "'                 " & vbCrLf
                    SQL = SQL & "     , ������������    = 5                                 " & vbCrLf '1 : óġ��, 5 : �Ϸ�
                    SQL = SQL & "     , ���̷ο�        = '" & strJudge & "'                " & vbCrLf
                    SQL = SQL & " Where �����          = '" & strYear & "'                 " & vbCrLf
                    SQL = SQL & "   and �����          = '" & strMonth & "'                " & vbCrLf
                    SQL = SQL & "   and ������          = '" & strDay & "'                  " & vbCrLf
                    SQL = SQL & "   and íƮ��ȣ        = '" & strChartNo & "'              " & vbCrLf
                    SQL = SQL & "   And ó���ڵ�        = '" & mGetP(strTestCd, 1, "|") & "'" & vbCrLf
                    SQL = SQL & "   And �����ڵ�        = '" & mGetP(strTestCd, 2, "|") & "'" & vbCrLf
                            
                    Call SetSQLData("�������", SQL, "A")
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SaveTransData_MEDICHART" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
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
    Dim Prm5            As New ADODB.Parameter
    Dim Prm6            As New ADODB.Parameter
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
    
    With frmMain
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
        
        '-- Local���� ȯ�ں��� ����� ��������
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
                
                '-- ���������
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                
                If strPatID <> "" And strTestCd <> "" And sResult <> "" Then
                    '-- ��������
                    Set AdoCmd = New ADODB.Command
                    Set AdoCmd.ActiveConnection = AdoCn
                    With AdoCmd
                        .CommandTimeout = 15
                        .CommandText = "sp_�˻簪����"
                        .CommandType = adCmdStoredProc
                        
                        Set Prm1 = .CreateParameter("receiptdate", adDate, adParamInput, 30, Format(strHospDate, "####-##-##"))
                        .Parameters.Append Prm1
                        
                        Set Prm2 = .CreateParameter("receiptnum", adVarChar, adParamInput, 30, strPatID)
                        .Parameters.Append Prm2
                        
                        Set Prm3 = .CreateParameter("labcode", adVarChar, adParamInput, 30, strTestCd)
                        .Parameters.Append Prm3
                        
                        Set prm4 = .CreateParameter("resultvalue", adVarChar, adParamInput, 4000, sResult)
                        .Parameters.Append prm4
                        
                        Set Prm5 = .CreateParameter("resultvalue2", adVarChar, adParamInput, 50, "")
                        .Parameters.Append Prm5
                        
                        Set Prm6 = .CreateParameter("resultvalue3", adVarChar, adParamInput, 50, "")
                        .Parameters.Append Prm6
                        
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
                        
                        Set prm14 = .CreateParameter("computerid", adVarChar, adParamInput, 30, CStr(frmMain.wSck.LocalIP))
                        .Parameters.Append prm14
                        
                        Set prm15 = .CreateParameter("overwrite", adVarChar, adParamInput, 1, "")
                        .Parameters.Append prm15
                        
                        Set prm16 = .CreateParameter("updatecount", adVarChar, adParamInputOutput, 100, 0)
                        .Parameters.Append prm16
    
                        .Execute
                        
                        strRet = .Parameters("updatecount").Value
                        If strRet = "0" Or strRet = "1" Then
                            '-- ���强��
                            frmMain.lblIFStatus.Caption = strPatID & " �˻��� ����"
                            Set AdoCmd = Nothing
                            SaveTransData_LABSPEAR = 1
                        Else
                            '-- �������
                            frmMain.lblIFStatus.Caption = strPatID & " �˻��� �������"
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
                        SQL = SQL & CStr(frmMain.wSck.LocalIP) & "," & vbCrLf
                        SQL = SQL & "" & "," & vbCrLf
                        SQL = SQL & "0" & "," & vbCrLf
                        SQL = SQL & "strRet:" & strRet & vbCrLf
                
                        Call SetSQLData("�������", SQL, "A")
                    End With
                End If
                RsLocal.MoveNext
            Loop
        End If
        
        '-- �кθ޸�����
        If strPatID <> "" And strCmnt <> "" And (strRet = "0" Or strRet = "1") Then
            Set AdoCmd = New ADODB.Command
            Set AdoCmd.ActiveConnection = AdoCn
            With AdoCmd
                .CommandTimeout = 15
                .CommandText = "sp_�˻��кθ޸�����"
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
                
                Set prmcmt6 = .CreateParameter("computerid", adVarChar, adParamInput, 30, CStr(frmMain.wSck.LocalIP))
                .Parameters.Append prmcmt6
                
                Set prmcmt7 = .CreateParameter("overwrite", adVarChar, adParamInput, 1, "")
                .Parameters.Append prmcmt7
                
                Set prmcmt8 = .CreateParameter("updatecount", adVarChar, adParamInputOutput, 100, 0)
                .Parameters.Append prmcmt8
                
                .Execute
                
                strRet = .Parameters("updatecount").Value
                If strRet = "0" Or strRet = "1" Then
                    '-- ���强��
                    Set AdoCmd = Nothing
                    SaveTransData_LABSPEAR = 1
                    
                Else
                    '-- �������
                    'MsgBox "�˻��� ������� " & .Parameters("updatecount").Value, vbInformation + vbOKOnly
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SaveTransData_LABSPEAR" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
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
    
    With frmMain
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
        
        '-- Local���� ȯ�ں��� ����� ��������
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
                
                '-- ���������
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                
                If strPatID <> "" And strTestCd <> "" And sResult <> "" Then
                    '-- ��������
                    SQL = ""
                    SQL = SQL & "Update LisiLib.Minterface                      " & vbCrLf
                    SQL = SQL & "   Set Result      = '" & Trim(sResult) & "'   " & vbCrLf
                    SQL = SQL & "     , Rltflag     = 'N'                       " & vbCrLf
                    SQL = SQL & "     , Updtdate    = (select substring(char(curdate()),1,4) || substring(char(curdate()),6,2) || substring(char(curdate()),9,2) || substring(char(curtime()),4,2) || substring(char(curtime()),7,2) || substring(char(curtime()),10,2) from sysibm.sysdummy1) " & vbCrLf
                    SQL = SQL & "     , Testercode  = '" & gHOSP.USERID & "'    " & vbCrLf
                    SQL = SQL & "     , Flag        = '2'                       " & vbCrLf
                    SQL = SQL & "     , Updtempl    = '" & gHOSP.USERID & "'    " & vbCrLf
                    '�ڸ�Ʈ
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
                    
                    
                    '�ڸ�Ʈ : frltcode = �ڵ�
                    '����   :
                    
                    Call SetSQLData("�������", SQL, "A")
                    AdoCn.Execute SQL
                    
                End If
                RsLocal.MoveNext
            Loop
        End If
        
        '-- �������� (��� ������ �Ϸ�Ǹ� �ش� procedure�� call �Ѵ�)
        'batch slrtrm55p(pmach : char(3) => ����ڵ�,
        '                perr  : char(1) => ����Ȯ�� �� �����ڵ�),
        '
        'real  slrtrm56p(pbarc : char(12) => ���ڵ��ȣ,
        '                pmach : char(3) => ����ڵ�,
        '                perr  : char(1) => ����Ȯ�� �� �����ڵ�)
        
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
                    
                Call SetSQLData("�������", "���ν���ȣ��", "A")
            
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SaveTransData_SCL" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
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
    
    With frmMain
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
        
        '-- Local���� ȯ�ں��� ����� ��������
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
                
                '-- ���������
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                  
                '-- �������� Ű ��������
                If strOrdCd = "" Then
                    strOrdCd = GetSampleSubITEM(strBarcode, strTestCd, strHospDate, strChartNo)
                End If
                
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" And strOrdCd <> "" Then
                    strDate = Format(Now, "yyyy-mm-dd")
                    strTime = Format(Now, "hh:mm:ss")
                
                    '-- ��������
                    SQL = ""
                    SQL = SQL & "UPDATE ME_LABDAT                           " & vbCrLf
                    SQL = SQL & "   Set LABRESULT = '" & sResult & "'       " & vbCrLf  '�˻���
                    SQL = SQL & "     , LABENDDEP = '2'                     " & vbCrLf  'ó������       2:����, 3:����Է�
                    SQL = SQL & "     , LABRSTDTE = '" & strDate & "'       " & vbCrLf  '����Է�����   YYYY-MM-DD
                    SQL = SQL & "     , LABRSTTIM = '" & strTime & "'       " & vbCrLf  '����Է�����   YYYY-MM-DD
                    SQL = SQL & "     , LABRSTUID = '" & gHOSP.USERID & "'  " & vbCrLf  '����Է�ID
                    SQL = SQL & "     , LABRSTCOM = '" & gHOSP.MACHNM & "'  " & vbCrLf  '����Է���ǻ�͸�
                    SQL = SQL & " WHERE LABATTEND = '" & strPatID & "'      " & vbCrLf  '������ȣ
                    SQL = SQL & "   And LABODRCOD = '" & strTestCd & "'     " & vbCrLf  '�˻��ڵ�
                    SQL = SQL & "   And LABODRSTP = '" & strOrdCd & "'      " & vbCrLf  '�˻��Ϸù�ȣ
                    SQL = SQL & "   And LABODRDTE = '" & strHospDate & "'   " & vbCrLf
'                    SQL = SQL & "   And LABBARCOD = '" & strBarcode & "'    " & vbCrLf  '���ڵ�
                    
                    Call SetSQLData("�������", SQL, "A")
                    AdoCn.Execute SQL
                                        
                    '-- ���º���
                    SQL = ""
                    SQL = SQL & "UPDATE ME_DAT                              " & vbCrLf
                    SQL = SQL & "   Set DATENDDEP   = '2'                   " & vbCrLf  'ó������       2:����, 3:����Է�
                    SQL = SQL & "     , DATRSTDTE = '" & strDate & "'       " & vbCrLf  '����Է�����   YYYY-MM-DD
                    SQL = SQL & "     , DATRSTTIM = '" & strTime & "'       " & vbCrLf  '����Է½ð�   hh:mm:ss
                    SQL = SQL & "     , DATRSTUID = '" & gHOSP.USERID & "'  " & vbCrLf  '����Է�ID
                    SQL = SQL & "     , DATRSTCOM = '" & gHOSP.MACHNM & "'  " & vbCrLf  '����Է���ǻ�͸�
                    SQL = SQL & " WHERE DATATTEND = '" & strPatID & "'      " & vbCrLf  '������ȣ
                    SQL = SQL & "   And DATODRCOD = '" & strTestCd & "'     " & vbCrLf  '�˻��ڵ�
                    SQL = SQL & "   And DATODRSTP = '" & strOrdCd & "'      " & vbCrLf  '�˻��Ϸù�ȣ
                    SQL = SQL & "   And DATODRDTE = '" & strHospDate & "'"
                    'SQL = SQL & "   And DATBARCOD = '" & strBarcode & "'    " & vbCrLf  '���ڵ�
                    
                    Call SetSQLData("���º���", SQL, "A")
                    AdoCn.Execute SQL
                    
                    blnSave = True
                            
                End If
                RsLocal.MoveNext
            Loop
        End If
        
        RsLocal.Close
        
'        If blnSave = True Then
'            '-- ���º���
'            SQL = ""
'            SQL = SQL & "UPDATE ME_DAT Set " & vbCrLf
'            SQL = SQL & "   Set DATENDDEP   = '2' " & vbCrLf         'ó������       2:����, 3:����Է�
'            SQL = SQL & "     , DATRSTDTE = '" & strDate & "', " & vbCrLf  '����Է�����   YYYY-MM-DD
'            SQL = SQL & "     , DATRSTTIM = '" & strTime & "', " & vbCrLf  '����Է½ð�   hh:mm:ss
'            SQL = SQL & "     , DATRSTUID = '" & gHOSP.USERID & "', " & vbCrLf  '����Է�ID
'            SQL = SQL & "     , DATRSTCOM = '" & gHOSP.MACHNM & "' " & vbCrLf    '����Է���ǻ�͸�
'            SQL = SQL & " WHERE DATATTEND = '" & strPatID & "'" & vbCrLf '������ȣ
'            SQL = SQL & "   And DATODRCOD = " & gAllOrdCd & vbCrLf     'ó���ڵ�
'            SQL = SQL & "   And DATODRSTP = '" & strOrdCd & "'"       '�˻��Ϸù�ȣ
'            SQL = SQL & "   And DATODRDTE = '" & strHospDate & "'"
'            SQL = SQL & "   And DATBARCOD = '" & strBarcode & "'" & vbCr  '���ڵ�
'
'            Call SetSQLData("���º���", "�������º��� : " & SQL)
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SaveTransData_BIT70" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
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
    
On Error GoTo ErrHandle
    
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""

    With frmMain
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
        
        
        '-- Local���� ȯ�ں��� ����� ��������
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
                
                '-- ���������
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                    '-- ��������
                    SQL = ""
                    SQL = SQL & "Exec UP_LIS_INTERFACE_U$001 " & dblBarno & "," & strTestCd & "," & sResult & "," & gHOSP.MACHCD
                        
                    Set AdoCmd = New ADODB.Command
                    Set AdoCmd.ActiveConnection = AdoCn
                    With AdoCmd
                        .CommandTimeout = 15
                        .CommandText = "UP_LIS_INTERFACE_U$001"
                        .CommandType = adCmdStoredProc
                        
                        Set Prm1 = .CreateParameter("BCODE_NO", adInteger, adParamInput, 30, dblBarno)      '���ڵ��ȣ
                        .Parameters.Append Prm1
    
                        Set Prm2 = .CreateParameter("ORD_CD", adVarChar, adParamInput, 10, strTestCd)       'ó���ڵ�
                        .Parameters.Append Prm2
    
                        Set Prm3 = .CreateParameter("RESULT_NM", adVarChar, adParamInput, 4000, sResult)    '�����
                        .Parameters.Append Prm3
    
                        Set prm4 = .CreateParameter("EQP_CD", adVarChar, adParamInput, 15, gHOSP.MACHCD)    '����ڵ�
                        .Parameters.Append prm4
    
                        .Execute
                        
                    End With
                    
                    Call SetSQLData("�������", SQL, "A")
                    AdoCn.Execute SQL
                    
                End If
                RsLocal.MoveNext
            Loop
        End If
        
        RsLocal.Close
        
        SaveTransData_MCC = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_MCC = -1
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "SaveTransData_MCC" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

Function SaveTransData_NU(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    Dim RS_L            As ADODB.Recordset
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
    Dim strAllResult    As String
    Dim sParam          As String
    Dim sRcvData        As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_NU = -1
        intRow = 0
        strJudge = ""
        blnSave = False
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        
        strTime = Format(Now, "hh:mm:ss")
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
'        If Trim(strPatID) = "" Then
'            Exit Function
'        End If
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMCODESUB,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '����ڵ�
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '�����ȣ
        SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '���ڵ�
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '�˻���
        
        'Call SetSQLData("���ð����ȸ", SQL)
        
        Set RS_L = New ADODB.Recordset
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            Do Until RS_L.EOF
                intRow = intRow + 1
                strEqpcd = RS_L.Fields("EQUIPCODE").Value & ""
                strOrdCd = RS_L.Fields("ORDERCODE").Value & ""
                strTestCd = RS_L.Fields("EXAMCODE").Value & ""
                strSubCode = RS_L.Fields("EXAMCODESUB").Value & ""
                sResult1 = RS_L.Fields("EQUIPRESULT").Value & ""
                sResult2 = RS_L.Fields("RESULT").Value & ""
                
                '-- ���������
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                
                strDate = Format(Now, "yyyymmddhhmmss")
                
                If Mid(strTestCd, 1, 3) = "LPD" Then
                    If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                        If gHOSP.BARUSE = "Y" Then
                            strAllResult = strAllResult & strTestCd & "" & sResult & "" & strDate & ""
                        Else
                            strAllResult = strAllResult & strTestCd & "" & sResult & "" & strDate & ""
                        End If
                    End If
                End If
                RS_L.MoveNext
            Loop
        End If
        
        RS_L.Close
        
        
        If strAllResult <> "" Then
                
            'http://emr016.cmcnu.or.kr/cmcnu/.live?submit_id=TXLII00101&business_id=li&
            '                                   ex_interface=10611884|016&
            '                                   bcno=E66YT0J70&
            '                                   result=LPD328Negative20200514LPD331Negative20200514LPD329Negative20200514&
            '                                   instcd=016&eqmtcd=P08&userid=10611884&
            
            'http://emr016.cmcnu.or.kr/cmcnu/.live?submit_id=TXLII00101&business_id=li&
            '                                   ex_interface=21903251|016&
            '                                   bcno=E66YV07J0&
            '                                   result=XXXXX21.6020200515120357LPD331Negative20200515120357
            '                                   &instcd=016&eqmtcd=P08&userid=21903251&

            sParam = ""
            sParam = sParam & "submit_id=TXLII00101&"
            sParam = sParam & "business_id=li&"
            sParam = sParam & "ex_interface=" & frmMain.txtTestID.Text & "|" & gHOSP.HOSPCD & "&"       '�����ID|����ڵ�
            sParam = sParam & "bcno=" & strBarcode & "&"                                    '��ü��ȣ(���ڵ�)
            sParam = sParam & "result=" & strAllResult & "&"                                '���
            sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"                                '����ڵ�
            sParam = sParam & "eqmtcd=" & gHOSP.MACHCD & "&"                                '����ڵ�
            sParam = sParam & "userid=" & frmMain.txtTestID.Text & "&"                                '�����ID
            
            sRcvData = OpenURLWithIE2(gHOSP.APIURL & sParam, frmMain.Inet1)

            Call SetSQLData("�������", "Param:" & gHOSP.APIURL & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine, "A")
            
            If InStr(1, sRcvData, "<?xml version") > 0 Then
                SaveTransData_NU = 1
            Else
                SaveTransData_NU = -1
            End If
            
        End If
        
        SaveTransData_NU = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_NU = -1
    
End Function

'Function SaveTransData_SUNGMO(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
'    Dim RS_L            As ADODB.Recordset
'    Dim intRow          As Integer
'    Dim strDate         As String
'    Dim strTime         As String
'
'    Dim strSaveSeq      As String
'    Dim strExamDate     As String
'    Dim strHospDate     As String
'    Dim strBarcode      As String
'    Dim strChartNo      As String
'    Dim strPatID        As String
'    Dim strIO           As String
'    Dim strKey1         As String
'    Dim strSex          As String
'    Dim strAge          As String
'
'    Dim strOrdCd        As String
'    Dim strTestCd       As String
'    Dim strSubCode      As String
'    Dim strEqpcd        As String
'    Dim sResult         As String
'    Dim sResult1        As String
'    Dim sResult2        As String
'    Dim strRefVal       As String
'    Dim strJudge        As String
'    Dim blnSave         As Boolean
'    Dim strSeqS         As String
'    Dim strAllResult    As String
'    Dim sParam          As String
'    Dim sRcvData        As String
'
'On Error GoTo ErrHandle
'
'    With frmMain
'        SaveTransData_SUNGMO = -1
'        intRow = 0
'        strJudge = ""
'        blnSave = False
'
'        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
'        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
'
'        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
'        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
'        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
'        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
'
'        strTime = Format(Now, "hh:mm:ss")
'
'        If Trim(strBarcode) = "" Then
'            Exit Function
'        End If
'
''        If Trim(strPatID) = "" Then
''            Exit Function
''        End If
'
'              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMCODESUB,EQUIPRESULT,RESULT " & vbCr
'        SQL = SQL & "  FROM PATRESULT " & vbCr
'        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '����ڵ�
'        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '�����ȣ
'        SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '���ڵ�
'        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '�˻���
'
'        'Call SetSQLData("���ð����ȸ", SQL)
'
'        Set RS_L = New ADODB.Recordset
'        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'            Do Until RS_L.EOF
'                intRow = intRow + 1
'                strEqpcd = RS_L.Fields("EQUIPCODE").Value & ""
'                strOrdCd = RS_L.Fields("ORDERCODE").Value & ""
'                strTestCd = RS_L.Fields("EXAMCODE").Value & ""
'                strSubCode = RS_L.Fields("EXAMCODESUB").Value & ""
'                sResult1 = RS_L.Fields("EQUIPRESULT").Value & ""
'                sResult2 = RS_L.Fields("RESULT").Value & ""
'
'                '-- ���������
'                If gHOSP.SAVELIS = "Y" Then
'                    sResult = sResult2
'                Else
'                    sResult = sResult1
'                End If
'
'                strDate = Format(Now, "yyyymmddhhmmss")
'
'                'gCFXCmnt.PB5
'
'                If Mid(strTestCd, 1, 3) = "LPD" Then
'                    If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
'                        If gHOSP.BARUSE = "Y" Then
'                            strAllResult = strAllResult & strTestCd & "" & sResult & "" & strDate & ""
'                        Else
'                            strAllResult = strAllResult & strTestCd & "" & sResult & "" & strDate & ""
'                        End If
'                    End If
'                End If
'                RS_L.MoveNext
'            Loop
'        End If
'
'        RS_L.Close
'
'
'        If strAllResult <> "" Then
'
'            'http://emr016.cmcnu.or.kr/cmcnu/.live?submit_id=TXLII00101&business_id=li&
'            '                                   ex_interface=10611884|016&
'            '                                   bcno=E66YT0J70&
'            '                                   result=LPD328Negative20200514LPD331Negative20200514LPD329Negative20200514&
'            '                                   instcd=016&eqmtcd=P08&userid=10611884&
'
'            'http://emr016.cmcnu.or.kr/cmcnu/.live?submit_id=TXLII00101&business_id=li&
'            '                                   ex_interface=21903251|016&
'            '                                   bcno=E66YV07J0&
'            '                                   result=XXXXX21.6020200515120357LPD331Negative20200515120357
'            '                                   &instcd=016&eqmtcd=P08&userid=21903251&
'
'            sParam = ""
'            sParam = sParam & "submit_id=TXLII00101&"
'            sParam = sParam & "business_id=li&"
'            sParam = sParam & "ex_interface=" & frmMain.txtTestID.Text & "|" & gHOSP.HOSPCD & "&"       '�����ID|����ڵ�
'            sParam = sParam & "bcno=" & strBarcode & "&"                                    '��ü��ȣ(���ڵ�)
'            sParam = sParam & "result=" & strAllResult & "&"                                '���
'            sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"                                '����ڵ�
'            sParam = sParam & "eqmtcd=" & gHOSP.MACHCD & "&"                                '����ڵ�
'            sParam = sParam & "userid=" & frmMain.txtTestID.Text & "&"                                '�����ID
'
'
'            '--�Ǵ� :  2018.09.21 ����
'            sParam = "http://" & gKNUH.IP & gKNUH.RSTURL & "&"
'            sParam = sParam & "Data1=" & lsID
'            sParam = sParam & "%03" & strEqpcd
'            sParam = sParam & "%17" & sResult
'            sParam = sParam & "%17" & "" 'sCommt
'            sParam = sParam & "%17" & Format(Now, "yyyymmddhhmmss")
'            sParam = sParam & "%17" & ""
'            If iRow = 1 Then
'                sParam = sParam & "%17" & sComment
'            Else
'                sParam = sParam & "%17"
'            End If
'            sParam = sParam & "&Data2=" & gEquipCode
'
'
'            sRcvData = OpenURLWithIE2(gHOSP.APIURL & sParam, frmMain.Inet1)
'
'            Call SetSQLData("�������", "Param:" & gHOSP.APIURL & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine, "A")
'
'            If InStr(1, sRcvData, "<?xml version") > 0 Then
'                SaveTransData_NU = 1
'            Else
'                SaveTransData_NU = -1
'            End If
'
'        End If
'
'        SaveTransData_NU = 1
'
'    End With
'
'Exit Function
'
'ErrHandle:
'    SaveTransData_NU = -1
'
'End Function


Function SaveTransData_KHNMC(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    Dim RS_L            As ADODB.Recordset
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
    Dim sComment        As String
    Dim strRefVal       As String
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    Dim strAllResult    As String
    Dim sParam          As String
    Dim sRcvData        As String
    
'On Error GoTo ErrHandle
On Error Resume Next

    With frmMain
        SaveTransData_KHNMC = -1
        intRow = 0
        strJudge = ""
        blnSave = False
        strAllResult = ""
        sComment = ""
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        
        
       ' strTime = Format(Now, "hh:mm:ss")
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
'        If Trim(strPatID) = "" Then
'            Exit Function
'        End If
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMCODESUB,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '����ڵ�
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '�����ȣ
        SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '���ڵ�
        'SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '�˻���
        
        Call SetSQLData("���ð����ȸ", SQL)
        
        
        Set RS_L = New ADODB.Recordset
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            Do Until RS_L.EOF
                intRow = intRow + 1
                strEqpcd = RS_L.Fields("EQUIPCODE").Value & ""
                strOrdCd = RS_L.Fields("ORDERCODE").Value & ""
                strTestCd = RS_L.Fields("EXAMCODE").Value & ""
                'strSubCode = RS_L.Fields("EXAMCODESUB").Value & ""
                sResult1 = RS_L.Fields("EQUIPRESULT").Value & ""
                sResult2 = RS_L.Fields("RESULT").Value & ""
                
                '-- ���������
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                
                '-- 2020-09-16 �߰�
                If sResult = "Inconclusive Result" Then
                    sResult = "Inconclusive"
                End If
                
                strDate = Format(Now, "yyyymmddhhmmss")
                
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                    'strAllResult = strAllResult & strOrdCd & "" & sResult & "" & strDate & ""
                    
                    strAllResult = strAllResult & strOrdCd & ""        'IF�ڵ�
                    strAllResult = strAllResult & sResult & ""         '�����
                    strAllResult = strAllResult & sComment & ""        '�ڸ�Ʈ
                    strAllResult = strAllResult & strDate & ""         '�˻��Ͻ�
                End If
                RS_L.MoveNext
            Loop
        End If
        
        RS_L.Close
        
        
        If strAllResult <> "" Then
            sParam = ""
            sParam = sParam & "&Data1=" & strBarcode & "" & strAllResult
            sParam = sParam & "&Data2=" & gHOSP.MACHCD
            
            sRcvData = OpenURLWithIE2("http://" & gHOSP.APIURL & "/lis/jindangeomsaweb/GyeolGwaIF.live?Mode=reqGetIFGyeolGwaDR&" & sParam, frmMain.Inet1)

            Call SetSQLData("�������", "Param>>" & "http://" & gHOSP.APIURL & "/lis/jindangeomsaweb/GyeolGwaIF.live?Mode=reqGetIFGyeolGwaDR&" & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine, "A")
            
            If InStr(1, sRcvData, "<?xml version") > 0 Then
                SaveTransData_KHNMC = 1
            Else
                SaveTransData_KHNMC = -1
            End If

            '�Ұ�����
            sParam = ""
            sParam = sParam & "&Data1=" & strPatID              '��Ϲ�ȣ
            sParam = sParam & "&Data2=" & strBarcode            '���ڵ��ȣ
            gCFXCmnt.SDCOVID = Replace(gCFXCmnt.SDCOVID, "CHR(10)CHR(13)", "%0D")
            sParam = sParam & "&Data3=" & gCFXCmnt.SDCOVID      '�ǻ�Ұ� ����
            sParam = sParam & "&Data4=" & gHOSP.MACHCD          '����ڵ�
            sParam = sParam & "&Data5=" & "91"                  '�����ڵ�
            
            sRcvData = OpenURLWithIE2("http://" & gHOSP.APIURL & "/lis/jindangeomsaweb/GyeolGwaIF.live?Mode=reqSetCLOBrstIF&" & sParam, frmMain.Inet1)

            Call SetSQLData("�Ұ�����", "Param>>" & "http://" & gHOSP.APIURL & "/lis/jindangeomsaweb/GyeolGwaIF.live?Mode=reqSetCLOBrstIF&" & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine, "A")
            
            If InStr(1, sRcvData, "<?xml version") > 0 Then
                SaveTransData_KHNMC = 1
            Else
                SaveTransData_KHNMC = -1
            End If
        End If
        
        SaveTransData_KHNMC = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_KHNMC = -1
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SaveTransData_KHNMC" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function


Function SaveTransData_HCMS(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
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
    Dim Prm5            As New ADODB.Parameter
    Dim Prm6            As New ADODB.Parameter
    Dim prm7            As New ADODB.Parameter
    
    Dim strDate         As String
    Dim strTime         As String
    
    
On Error GoTo ErrHandle
    
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""
    
    strDate = Format(Now, "yyyymmdd")
    strTime = Format(Now, "hhmmss")

    With frmMain
        SaveTransData_HCMS = -1
        
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
        
        
        '-- Local���� ȯ�ں��� ����� ��������
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EQUIPRESULT,RESULT,REFJUDGE    " & vbCrLf
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
                sResult1 = RsLocal.Fields("EQUIPRESULT").Value & ""
                sResult2 = RsLocal.Fields("RESULT").Value & ""
                strJudge = RsLocal.Fields("REFJUDGE").Value & ""
                
                '-- ���������
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                    '-- ��������
                    SQL = ""
                    SQL = SQL & "Exec LIME.PR_LS_PRC_LAB_MAC_RSLT " & sResult & "," & strDate & "," & strTime & "," & gHOSP.USERID & "," & strBarcode & "," & strTestCd
                         
'                    Set AdoCmd = New ADODB.Command
'                    Set AdoCmd.ActiveConnection = AdoCn
'                    With AdoCmd
'                        .CommandTimeout = 15
'                        .CommandText = "LIME.PR_LS_PRC_LAB_MAC_RSLT"
'                        .CommandType = adCmdStoredProc
'
'                        Set Prm1 = .CreateParameter("IV_RSLT_VALUE", adInteger, adParamInput, 30, dblBarno)      '���ڵ��ȣ
'                        .Parameters.Append Prm1
'
'                        Set Prm2 = .CreateParameter("IV_RSLT_DD", adVarChar, adParamInput, 10, strTestCd)       'ó���ڵ�
'                        .Parameters.Append Prm2
'
'                        Set Prm3 = .CreateParameter("RESULT_NM", adVarChar, adParamInput, 4000, sResult)    '�����
'                        .Parameters.Append Prm3
'
'                        Set prm4 = .CreateParameter("EQP_CD", adVarChar, adParamInput, 15, gHOSP.MACHCD)    '����ڵ�
'                        .Parameters.Append prm4
'
'                        .Execute
'
'                    End With
                    
                    With AdoCmd
                        .CommandType = adCmdStoredProc
                        .Prepared = True
                        .CommandTimeout = 15
                        .CommandText = "LIME.PR_LS_PRC_LAB_MAC_RSLT"
                        .CommandType = adCmdStoredProc
                        .Parameters.Refresh
                        .Parameters("IV_RSLT_VALUE").Value = sResult            '�˻���
                        .Parameters("IV_RSLT_DD").Value = strDate               '�Է�����
                        .Parameters("IV_RSLT_TIME").Value = strTime             '�Է½ð�
                        .Parameters("IV_RSLT_INPS_ID").Value = gHOSP.USERID     '����Է���
                        .Parameters("IV_SPCM_RCPN_NO").Value = strBarcode       '���ڵ�
                        .Parameters("IV_EXMN_CD").Value = strTestCd             '�����ڵ�
                       
                        Set AdoRs = AdoCmd.Execute
                        Set AdoRs = Nothing
                    End With
                    
                    
                    Call SetSQLData("�������", SQL, "A")
                    AdoCn.Execute SQL
                    
                End If
                RsLocal.MoveNext
            Loop
        End If
        
        RsLocal.Close
        
        SaveTransData_HCMS = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_HCMS = -1
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_SaveTransData_HCMS" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
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

    With frmMain
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
        
        
        '-- Local���� ȯ�ں��� ����� ��������
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
                
                '-- ���������
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                    '-- ��������
                    SQL = ""
                    SQL = SQL & "Update TW_HSP_OCS.TWEXAM_RESULTC           " & vbCrLf
                    SQL = SQL & "   Set STATUS      = '4'                   " & vbCrLf  '�˻����
                    SQL = SQL & "     , RESULT      = '" & sResult & "'     " & vbCrLf  '�˻���
                    SQL = SQL & "     , RESULTDATE  = TRUNC(SYSDATE)        " & vbCrLf  '�˻����۽ð�
                    SQL = SQL & " Where SPECNO      = '" & strBarcode & "'  " & vbCrLf  '��ü��ȣ
'                    SQL = SQL & "   And MASTERCODE  = 'LH1P01'    " & vbCrLf  '�������ڵ� LH1P01
                    SQL = SQL & "   And SUBCODE     = '" & strTestCd & "'   " & vbCrLf  '�˻��ڵ�
                    SQL = SQL & "   And STATUS      <= '3'                  " & vbCrLf  '�˻����(=��ü����)
                    
                    Call SetSQLData("�������", SQL, "A")
                    AdoCn.Execute SQL
                
                    '-- ���¾�����Ʈ
                    SQL = ""
                    SQL = SQL & "Update TW_HSP_OCS.TWEXAM_SPECMST           " & vbCrLf
                    SQL = SQL & "   Set STATUS     = '3'                    " & vbCrLf '�˻���� [������(3:�����Ȯ��, 4:�κ�����)]
                    SQL = SQL & "     , RESULTDATE = TRUNC(SYSDATE)         " & vbCrLf
                    SQL = SQL & " Where SPECNO     = '" & strBarcode & "'   " & vbCrLf '��ü��ȣ
                    SQL = SQL & "   And STATUS     <= '3'                   " & vbCrLf '�˻���� [3:��ü����]
                    
                    Call SetSQLData("��������", SQL, "A")
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "SaveTransData_TWIN" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function


'-- �˻縶���� ��ȸ
Public Sub GetTestList()
    Dim intRow          As Long

    intRow = 0
    gAllTestCd = ""
    Erase gArrEQP

    SQL = ""
    SQL = SQL & "SELECT "
    SQL = SQL & "  SEQNO,SENDCHANNEL,RSLTCHANNEL,TESTCODE,TESTNAME,ABBRNAME " & vbCrLf
    SQL = SQL & " ,RESPRECUSE,RESPREC,REFMLOW,REFMHIGH,REFFLOW,REFFHIGH,RESTYPE     " & vbCrLf
    SQL = SQL & " ,AMRLimit1,AMRLimit2,AMRLimit3,AMRLimit4,AMRLimit5        " & vbCrLf
    SQL = SQL & " ,AMRResult1,AMRResult2,AMRResult3,AMRResult4,AMRResult5   " & vbCrLf
    SQL = SQL & " ,AMRINResult                                              " & vbCrLf
    SQL = SQL & "  FROM EQPMASTER                                           " & vbCrLf
    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "'                    " & vbCrLf
    '�˻�ó���� ���� ����ϰ�� ó���ɷ� �������� �ϱ� ���ؼ�...
    SQL = SQL & " ORDER BY SEQNO ASC, TESTCODE DESC, TESTNAME "

    '-- Record Count ������
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        'ó���ڵ�, SUB�ڵ�� �߰� 16,17
        ReDim Preserve gArrEQP(AdoRs_Local.RecordCount, 17)

        Do Until AdoRs_Local.EOF
            intRow = intRow + 1
            gArrEQP(intRow, 1) = AdoRs_Local.Fields("SEQNO").Value & ""
            gArrEQP(intRow, 2) = AdoRs_Local.Fields("TESTCODE").Value & ""
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
            gArrEQP(intRow, 16) = ""    'ó���ڵ�� ���
            gArrEQP(intRow, 17) = ""    'SUB�ڵ�� ���

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

'-- �˻縶���͸� ��ȸ
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

    '-- Record Count ������
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


'-- �˻縶���� ��ȸ
Public Sub GetTestMaster(ByVal SPD As Object)
    Dim intRow          As Long

    SPD.MaxRows = 0
    intRow = 0

    SQL = ""
    SQL = SQL & "SELECT * " & vbCr
    SQL = SQL & "  FROM EQPMASTER " & vbCr
    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "'" & vbCr
    SQL = SQL & " ORDER BY SEQNO "

    '-- Record Count ������
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
                Call SetText(SPD, AdoRs_Local.Fields("RESTYPE").Value & "", intRow, colRESTYPE)
                
                Call SetText(SPD, AdoRs_Local.Fields("AMRLimit1").Value & "|" & AdoRs_Local.Fields("AMRResult1").Value & "", intRow, colRESTYPE + 1)
                Call SetText(SPD, AdoRs_Local.Fields("AMRLimit2").Value & "|" & AdoRs_Local.Fields("AMRResult2").Value & "", intRow, colRESTYPE + 2)
                Call SetText(SPD, AdoRs_Local.Fields("AMRLimit3").Value & "|" & AdoRs_Local.Fields("AMRResult3").Value & "", intRow, colRESTYPE + 3)
                Call SetText(SPD, AdoRs_Local.Fields("AMRLimit4").Value & "|" & AdoRs_Local.Fields("AMRResult4").Value & "", intRow, colRESTYPE + 4)
                Call SetText(SPD, AdoRs_Local.Fields("AMRLimit5").Value & "|" & AdoRs_Local.Fields("AMRResult5").Value & "", intRow, colRESTYPE + 5)
                Call SetText(SPD, AdoRs_Local.Fields("AMRLimit6").Value & "|" & AdoRs_Local.Fields("AMRResult6").Value & "", intRow, colRESTYPE + 6)
                Call SetText(SPD, AdoRs_Local.Fields("AMRLimit7").Value & "|" & AdoRs_Local.Fields("AMRResult7").Value & "", intRow, colRESTYPE + 7)
                Call SetText(SPD, AdoRs_Local.Fields("AMRINResult").Value & "", intRow, colRESTYPE + 8)
                
                AdoRs_Local.MoveNext
            Loop
            .RowHeight(-1) = 15
        End With
    End If

End Sub


'-- AMR������ ��ȸ
Public Sub GetAMRMaster(ByVal pSeqNo As Integer, ByVal pRCd As String, ByVal pTestCd As String)

    SQL = ""
    SQL = SQL & "SELECT * " & vbCrLf
    SQL = SQL & "  FROM AMRMASTER " & vbCr
    SQL = SQL & " WHERE EQUIPCD   = '" & gHOSP.MACHCD & "'" & vbCrLf
    SQL = SQL & "   AND SEQNO     = " & pSeqNo & vbCrLf
    SQL = SQL & "   AND RSLTCHANNEL  = '" & pRCd & "'" & vbCrLf
    SQL = SQL & "   AND TESTCODE     = '" & pTestCd & "'" & vbCrLf

    '-- Record Count ������
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            '������
            frmTestSet.txtAMRLimit(7).Text = AdoRs_Local.Fields("AMRLIMIT8").Value & ""
            frmTestSet.txtAMRLimit(8).Text = AdoRs_Local.Fields("AMRLIMIT9").Value & ""
            frmTestSet.txtAMRLimit(9).Text = AdoRs_Local.Fields("AMRLIMIT10").Value & ""
            frmTestSet.txtAMRLimit(10).Text = AdoRs_Local.Fields("AMRLIMIT11").Value & ""
            frmTestSet.txtAMRLimit(11).Text = AdoRs_Local.Fields("AMRLIMIT12").Value & ""
            frmTestSet.txtAMRLimit(12).Text = AdoRs_Local.Fields("AMRLIMIT13").Value & ""
            frmTestSet.txtAMRLimit(13).Text = AdoRs_Local.Fields("AMRLIMIT14").Value & ""
            
            frmTestSet.txtAMRResult(7).Text = AdoRs_Local.Fields("AMRRESULT8").Value & ""
            frmTestSet.txtAMRResult(8).Text = AdoRs_Local.Fields("AMRRESULT9").Value & ""
            frmTestSet.txtAMRResult(9).Text = AdoRs_Local.Fields("AMRRESULT10").Value & ""
            frmTestSet.txtAMRResult(10).Text = AdoRs_Local.Fields("AMRRESULT11").Value & ""
            frmTestSet.txtAMRResult(11).Text = AdoRs_Local.Fields("AMRRESULT12").Value & ""
            frmTestSet.txtAMRResult(12).Text = AdoRs_Local.Fields("AMRRESULT13").Value & ""
            frmTestSet.txtAMRResult(13).Text = AdoRs_Local.Fields("AMRRESULT14").Value & ""
            
            '��ġ��
            frmTestSet.txtAMRLimit(14).Text = AdoRs_Local.Fields("AMRLIMIT15").Value & ""
            frmTestSet.txtAMRLimit(15).Text = AdoRs_Local.Fields("AMRLIMIT16").Value & ""
            frmTestSet.txtAMRLimit(16).Text = AdoRs_Local.Fields("AMRLIMIT17").Value & ""
            frmTestSet.txtAMRLimit(17).Text = AdoRs_Local.Fields("AMRLIMIT18").Value & ""
            frmTestSet.txtAMRLimit(18).Text = AdoRs_Local.Fields("AMRLIMIT19").Value & ""
            
            frmTestSet.txtAMRResult(14).Text = AdoRs_Local.Fields("AMRRESULT15").Value & ""
            frmTestSet.txtAMRResult(15).Text = AdoRs_Local.Fields("AMRRESULT16").Value & ""
            frmTestSet.txtAMRResult(16).Text = AdoRs_Local.Fields("AMRRESULT17").Value & ""
            frmTestSet.txtAMRResult(17).Text = AdoRs_Local.Fields("AMRRESULT18").Value & ""
            frmTestSet.txtAMRResult(18).Text = AdoRs_Local.Fields("AMRRESULT19").Value & ""
            
            AdoRs_Local.MoveNext
        Loop
    End If

End Sub
''-- �˻���������� ��ȸ
'Public Sub GetOrderMST()
'    Dim intRow          As Long
'
''    gAllOrdCd = ""
''    intRow = 0
''
''    SQL = ""
''    SQL = SQL & "SELECT ORDERCODE FROM ORDMASTER " & vbCr
''    SQL = SQL & " ORDER BY ORDERCODE "
''
''    '-- Record Count ������
''    AdoCn_Local.CursorLocation = adUseClient
''    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
''    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
''        With frmMain.spdOrdMst
''            .MaxRows = AdoRs_Local.RecordCount
''            Do Until AdoRs_Local.EOF
''                intRow = intRow + 1
''                Call SetText(frmMain.spdOrdMst, AdoRs_Local.Fields("ORDERCODE").Value & "", intRow, 1)
''
''                If Trim(AdoRs_Local.Fields("ORDERCODE").Value) <> "" Then
''                    If intRow = 1 Or gAllTestCd = "" Then
''                        gAllOrdCd = "'" & AdoRs_Local.Fields("ORDERCODE").Value & "'"
''                    Else
''                        gAllOrdCd = gAllOrdCd & ",'" & AdoRs_Local.Fields("ORDERCODE").Value & "'"
''                    End If
''                End If
''
''                AdoRs_Local.MoveNext
''            Loop
''            .RowHeight(-1) = 12
''        End With
''    End If
'End Sub
'

'-- �˻��ڵ�� �˻�� ��ȸ
Public Function GetTestNm(ByVal pItem As String, Optional pFull As Boolean) As String
    Dim intRow          As Long

    GetTestNm = ""

    If pFull = True Then
        SQL = ""
        SQL = SQL & "SELECT TESTNAME AS ITEMNM FROM EQPMASTER " & vbCr
        SQL = SQL & " WHERE TESTCODE = '" & pItem & "'"
    Else
        SQL = ""
        SQL = SQL & "SELECT ABBRNAME AS ITEMNM FROM EQPMASTER " & vbCr
        SQL = SQL & " WHERE TESTCODE = '" & pItem & "'"
    End If

    '-- Record Count ������
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

'-- �˻��ڵ�� �˻�� ��ȸ
Public Function GetTestNm_ORD(ByVal pItem As String, Optional pFull As Boolean) As String
    Dim intRow          As Long

    GetTestNm_ORD = ""

    If pFull = True Then
        SQL = ""
        SQL = SQL & "SELECT TESTNAME AS ITEMNM FROM EQPMASTER " & vbCr
        SQL = SQL & " WHERE SENDCHANNEL = '" & pItem & "'"
    Else
        SQL = ""
        SQL = SQL & "SELECT ABBRNAME AS ITEMNM FROM EQPMASTER " & vbCr
        SQL = SQL & " WHERE SENDCHANNEL = '" & pItem & "'"
    End If

    '-- Record Count ������
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            GetTestNm_ORD = AdoRs_Local.Fields("ITEMNM").Value & ""
            AdoRs_Local.MoveNext
        Loop
    End If

    AdoRs_Local.Close

End Function


'-- �˻��ڵ�� �˻�� ��ȸ
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

    '-- Record Count ������
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
''-- �˻������ ���ä�� ��ȸ
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
'    '-- Record Count ������
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
''-- �˻��׸� ��ȸ
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
'    '-- Record Count ������
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
'                strErrMsg = "��    ġ : " & gHOSP.MACHNM & "GetTest" & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
'    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
'    frmErrMsg.txtErr = vbNewLine & strErrMsg
'    frmErrMsg.Show 'vbModal
'
'    Screen.MousePointer = 0
'
'End Function
'
''-- ��ũ����Ʈ ��ȸ
Public Sub GetWorkList(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As Object)

    Select Case gEMR
        Case "KHNMC"
                Call GetWorkList_KHNMC(pFrom, pTo, SPD)
        
        Case "NU"
                Call GetWorkList_NU(pFrom, pTo, SPD)
        
        Case "HCMS"
                Call GetWorkList_HCMS(pFrom, pTo, SPD)
        
        Case "JWINFO"
                Call GetWorkList_JWINFO(pFrom, pTo, SPD)
        
        Case "AMIS"                                 '�ƹ̽� ��ũ���
                Call GetWorkList_AMIS(pFrom, pTo, SPD)
        
        Case "EONM"                                 '�̿¿�
                Call GetWorkList_EONM(pFrom, pTo, SPD)

        Case "BIT70"                                '��Ʈ
                Call GetWorkList_BIT70(pFrom, pTo, SPD)

        Case "LABSPEAR"                             '�̳뺣��Ʈ(���Ƿ����)
                Call GetWorkList_LABSPEAR(pFrom, pTo, SPD)

        Case "SANSOFT"                              '�׽�Ʈ����
                Call GetWorkList_LABSPEAR(pFrom, pTo, SPD)

        Case "MEDICHART"                            '�޵�íƮ
                Call GetWorkList_MEDICHART(pFrom, pTo, SPD)

        Case "KCHART"                               '�ٴ����Ʈ
                Call GetWorkList_KCHART(pFrom, pTo, SPD)

'        Case "PHILL"
'                Call GetWorkList_PHILL(pFrom, pTo, SPD)
'
'        Case "MSINFOTEC"                    'MS������
'                Call GetWorkList_MSINFOTEC(pFrom, pTo, SPD)
'
'        Case "HANARO"                       '�ϳ����Ƿ����
'                Call GetWorkList_HANARO(pFrom, pTo, SPD)

'        Case "AMIS"                         '�ƹ̽�
'                Call GetWorkList_AMIS(pFrom, pTo, SPD)
'
'        Case "BIGUBCARE"
'                Call GetWorkList_BIGUBCARE(pFrom, pTo, SPD)
'
'        Case "BIT"                          '��Ʈ
'                Call GetWorkList_BIT(pFrom, pTo, SPD)
'
'        Case "BIT70"                        '��Ʈ HIB70
'                Call GetWorkList_BIT70(pFrom, pTo, SPD)
'
'        Case "EMEDI"                        '�̸޵�
'                Call GetWorkList_AMIS(pFrom, pTo, SPD)
'
'        Case "EASYS"                        '������, MCC
'                Call GetWorkList_EASYS(pFrom, pTo, SPD)
'
''        Case "EONM"                         '�̿¿�
''                Call GetWorkList_EONM(pFrom, pTo, SPD)
'
'        Case "GINUS"                         '������
''                Call GetWorkList_GINUS(pFrom, pTo, SPD)
'
'        Case "GSEN"                         '����Ŀ�´����̼���(��íƮ)
'                Call GetWorkList_MSINFOTEC(pFrom, pTo, SPD)
'
'        Case "HWASAN"                       'ȭ��
'                Call GetWorkList_HWASAN(pFrom, pTo, SPD)
'
'        Case "JAINCOM"                      '������
'                Call GetWorkList_JAINCOM(pFrom, pTo, SPD)
'
'        Case "JWINFO"                       '�߿�����
'                Call GetWorkList_JWINFO(pFrom, pTo, SPD)
'
'
'        Case "KOMAIN"                       '�߿�����
'                Call GetWorkList_KOMAIN(pFrom, pTo, SPD)
'
'        Case "KYU"                          '�Ǿ���б����� - ��ũ����Ʈ ��ɾ���
'                'Call GetWorkList_KYU(pFrom, pTo, SPD)
'

'        Case "MEDIIT"                       '�޵�IT(���Ƿ����)
'                Call GetWorkList_MEDIIT(pFrom, pTo, SPD)
'
'        Case "MEDITOLISS"                   '�Ƹ�����
'                Call GetWorkList_MEDITOLISS(pFrom, pTo, SPD)
'
'        Case "MCC"                          'MCC SP����
'                Call GetWorkList_MCC(pFrom, pTo, SPD)
'
'        Case "MOD"                          'MOD �ý���
'                Call GetWorkList_MOD(pFrom, pTo, SPD)
'
'        Case "MSINFOTEC"                    'MS������
'                Call GetWorkList_MSINFOTEC(pFrom, pTo, SPD)
'
'        Case "NEOSOFT"                      '�׿�����Ʈ
'                Call GetWorkList_NEOSOFT(pFrom, pTo, SPD)
'
'        Case "ONITGUM"                      '�¾�Ƽ ����
'                Call GetWorkList_ONITGUM(pFrom, pTo, SPD)
'
'        Case "ONITEMR"                      '�¾�Ƽ EMR
'                Call GetWorkList_ONITEMR(pFrom, pTo, SPD)
'
'        Case "PLIS"                         '���̽� ������ó
'                Call GetWorkList_PLIS(pFrom, pTo, SPD)
'
'        Case "SY"                           'SY
'                Call GetWorkList_SY(Format(pFrom, "yyyy-mm-dd"), Format(pTo, "yyyy-mm-dd"), SPD)
'
'        Case "TWIN"                         '��������
'                Call GetWorkList_TWIN(pFrom, pTo, SPD)
'
'        Case "UBCARE"                       '�ǻ��
'                Call GetWorkList_UBCARE(pFrom, pTo, SPD)

'        Case "WELL"                         '��Ŀ�ӽ�
'                Call GetWorkList_WELL(pFrom, pTo, SPD)

'        Case "ONIT"
'            Call GetWorkList_onit(pFrom, pTo, SPD)

'        Case "PLIS"
'            Call GetWorkList_PLIS(pFrom, pTo, SPD)
        Case Else


    End Select

End Sub

'���Ƿ���� OLD ����
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

    Call SetSQLData("��ũ��ȸ", SQL, "")

    '-- Record Count ������
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
        frmMain.lblComStatus.Caption = "��ũ����Ʈ ��ȸ ����ڰ� �����ϴ�."
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_GetWorkList_PHILL" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show vbModal

End Sub



Function GetOrderSeqCode(argExamDt As String, argPID As String, argPCD As String) As String
    Dim RS As ADODB.Recordset
    
    '-- SEQ ��������
    
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

    Call SetSQLData("SEQã��", SQL)

    '-- Record Count ������
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
    SQL = SQL & "   AND R.OKFL  <> 'Y'                                  " & vbCrLf   '-- ���Ȯ������ (Y / N)
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

    Call SetSQLData("��ũ��ȸ", SQL, "")

    '-- Record Count ������
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
        frmMain.lblComStatus.Caption = "��ũ����Ʈ ��ȸ ����ڰ� �����ϴ�."
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_GetWorkList_MSINFOTEC" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
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
    SQL = SQL & "   AND O.H141_NOTYYN   IN ('N','T')                            " & vbCrLf '������:T
    SQL = SQL & "   And O.H141_SUGACD   IN (" & gAllTestCd & ")                 " & vbCrLf
    SQL = SQL & " Group By O.H141_ODRDAT                                        " & vbCrLf
    SQL = SQL & "        , O.H141_TSAMPLENO                                     " & vbCrLf
    SQL = SQL & "        , O.H141_SEQNO                                         " & vbCrLf
    SQL = SQL & "        , P.A110_CHARTNO                                       " & vbCrLf
    SQL = SQL & "        , P.A110_PATNM                                         " & vbCrLf
    SQL = SQL & "        , P.A110_JUMIN1                                        " & vbCrLf
    SQL = SQL & "        , P.A110_SEX                                           " & vbCrLf
    SQL = SQL & " Order By O.H141_ODRDAT, O.H141_SEQNO                          " & vbCrLf

    Call SetSQLData("��ũ��ȸ", SQL, "")

    '-- Record Count ������
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
        frmMain.lblComStatus.Caption = "��ũ����Ʈ ��ȸ ����ڰ� �����ϴ�."
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_GetWorkList_EONM" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
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
    SQL = SQL & "   AND O.CLAS          = 4                                 " & vbCrLf '�ӻ󺴸�
    SQL = SQL & "   AND R.RESULTITEMCODE IN (" & gAllTestCd & ")            " & vbCrLf
    SQL = SQL & "   AND (R.NUMRESULTVAL = '' OR R.NUMRESULTVAL IS NULL)     " & vbCrLf
    SQL = SQL & " GROUP BY O.ACPTDATE, R.SPCMNO, P.PATID, P.PATNAME, P.SEX  " & vbCrLf
    SQL = SQL & " ORDER BY O.ACPTDATE, R.SPCMNO                             " & vbCrLf

    Call SetSQLData("��ũ��ȸ", SQL, "")

    '-- Record Count ������
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
        frmMain.lblComStatus.Caption = "��ũ����Ʈ ��ȸ ����ڰ� �����ϴ�."
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_GetWorkList_AMIS" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
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
    
    Call SetSQLData("��ũ��ȸ", SQL, "")

    '-- Record Count ������
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
        frmMain.lblComStatus.Caption = "��ũ����Ʈ ��ȸ ����ڰ� �����ϴ�."
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_GetWorkList_JWINFO" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    
    frmErrMsg.Show vbModal

End Sub

Public Sub GetWorkList_HCMS(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As Object)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean

    Dim i           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim strTestCds  As String
    Dim strWhere    As String
    
On Error GoTo ErrHandle

    Screen.MousePointer = 11
    blnSame = False
    strTestCds = ""
    strWhere = ""
    
    With frmMain
        If .optStatus(3).Value = True Then
            strWhere = ""
        Else
            If .optStatus(0).Value = True Then
                strWhere = "32"
            End If
            If .optStatus(1).Value = True Then
                strWhere = "24"
            End If
            If .optStatus(2).Value = True Then
                strWhere = "12"
            End If
        End If
    End With
    
    SQL = ""
    SQL = SQL & "SELECT SUPT_DEPT_RCPN_DT                               AS HOSPDATE " & vbCrLf
    SQL = SQL & "     , PID                                             AS PID      " & vbCrLf
    SQL = SQL & "     , PT_NM                                           AS PNAME    " & vbCrLf
    SQL = SQL & "     , SPCM_RCPN_NO                                    AS BARCODE  " & vbCrLf
    SQL = SQL & "     , DEPT_HNM                                        AS DEPTNM   " & vbCrLf
    SQL = SQL & "     , COUNT(IV_EXMN_CD)                               AS CNT      " & vbCrLf
    SQL = SQL & "  FROM LIME.VW_READ_LSMACIF                                        " & vbCrLf
    SQL = SQL & " WHERE SUPT_DEPT_RCPN_DT BETWEEN '" & pFrom & "' AND '" & pTo & "' " & vbCrLf
    If strWhere <> "" Then
        SQL = SQL & "   AND STATUS = '" & strWhere & "'                             " & vbCrLf
    End If
    SQL = SQL & "   AND LAB_MAC_CODE = '" & gHOSP.PARTCD & "'                       " & vbCrLf
    SQL = SQL & "   AND IV_EXMN_CD   IN (" & gAllTestCd & ")                        " & vbCrLf
    SQL = SQL & " GROUP BY SUPT_DEPT_RCPN_DT, PID, PT_NM, SPCM_RCPN_NO, DEPT_HNM    " & vbCrLf
    SQL = SQL & " ORDER BY SUPT_DEPT_RCPN_DT "
    
    Call SetSQLData("��ũ��ȸ", SQL, "")

    '-- Record Count ������
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
                    SetText SPD, Trim(RS.Fields("DEPTNM")) & "", intRow, colDEPT
                    SetText SPD, Trim(RS.Fields("CNT")) & "", intRow, colOCNT
                    
                    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    
                End If
                
            End With

            blnSame = False

            DoEvents

            RS.MoveNext
        Loop
    Else
        frmMain.lblComStatus.Caption = "��ũ����Ʈ ��ȸ ����ڰ� �����ϴ�."
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_GetWorkList_HCMS" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    
    frmErrMsg.Show vbModal

End Sub

Public Sub GetWorkList_NU(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As Object)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim J           As Integer
    Dim k           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim sParam      As String
    Dim strTestCds  As String
    Dim sRcvData    As String
    Dim varRcvData  As Variant
    Dim varTstCode  As Variant
    Dim strNames    As String
    Dim strXmlName  As String
    Dim strWorkNo   As String
    
    Dim l As Integer
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    SPD.MaxRows = 0
    
    blnSame = False
    strNames = ""
    l = 0
    
ReSearch:
    
    strTestCds = gHOSP.ORDCODE

    '��õ���𺴿�
    'http://emr016.cmcnu.or.kr/cmcnu/.live?submit_id=TRLII00101&business_id=li&instcd=016&eqmtcd=P08&startdd=20200514&enddd=20200514&

    sParam = ""
    sParam = sParam & "submit_id=TRLII00101&"                               'submit ID
    sParam = sParam & "business_id=li&"                                     'business_id
    sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"                        '����ڵ�
    sParam = sParam & "startdd=" & pFrom & "&"                              '�����۾�����
    sParam = sParam & "enddd=" & pTo & "&"                                  '�����۾�����
    sParam = sParam & "eqmtcd=" & gHOSP.MACHCD & "&"                        '����ڵ�
    
    sRcvData = OpenURLWithIE2(gHOSP.APIURL & sParam, frmMain.Inet1)
    
    Call SetSQLData("��ũ��ȸ", "Param:" & gHOSP.APIURL & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine)

    If InStr(1, sRcvData, "<?xml version") > 0 Then
        varRcvData = Split(sRcvData, "<worklist>")
    End If

    strXmlName = gHOSP.MACHNM & "_" & Format(CDate(Now), "yyyymmdd") & ".xml"

    Call SetXMLData(strXmlName, sRcvData)

    Call DisplayNode_InfoS(App.PATH & "\Xml\" & strXmlName, UBound(varRcvData))
    'Call DisplayNode_InfoS("D:\������Ʈ\VB\__NU\��õ���𺴿�_NU\����\worklist.xml", 10)

    Kill App.PATH & "\Xml\" & strXmlName
    
    If UBound(varRcvData) >= 1 Then
        For i = 0 To UBound(varRcvData) - 1
        'For i = 0 To 10
            With SPD
                .ReDraw = False
                blnSame = False
                
                '2019-12-11 �޸�
                '   rsltstat �� ó�� ���� �� ��...
                '   <rsltstat><![CDATA[-]]></rsltstat>
                '   <rsltstat><![CDATA[4]]></rsltstat>
                
                'If GetSampleTest_HDINFO(XmlSelectS.BCNO(i)) > 0 Then
                    
                If InStr(UCase(XmlSelectS.IFREQCDLIST(i)), "COVID19") > 0 Then
                    For J = 1 To SPD.DataRowCnt
                        strHospDate = GetText(SPD, J, colHOSPDATE)
                        strBarcode = GetText(SPD, J, colBARCODE)
                        If XmlSelectS.PRCPDD(i) & "" = strHospDate And XmlSelectS.BCNO(i) = strBarcode Then
                            blnSame = True
                            strNames = GetText(SPD, intRow, colITEMS)
                            strNames = strNames & "|" & GetTestNm(XmlSelectS.TestCd(i))
                            
                            SetText SPD, strNames, intRow, colITEMS
                            strNames = ""
                        End If
                    Next
                    
                    If blnSame = False Then
                        .MaxRows = .MaxRows + 1
                        intRow = .MaxRows
                
                        SetText SPD, "1", intRow, colCHECKBOX
                        SetText SPD, CStr(intRow), intRow, colSEQNO
                        SetText SPD, XmlSelectS.SPCACPTDT(i), intRow, colHOSPDATE
                        SetText SPD, XmlSelectS.BCNO(i), intRow, colBARCODE
                        SetText SPD, XmlSelectS.PID(i), intRow, colPID
                        SetText SPD, XmlSelectS.PATNM(i), intRow, colPNAME
                        SetText SPD, XmlSelectS.SEX(i), intRow, colPSEX
                        SetText SPD, XmlSelectS.AGE(i), intRow, colPAGE
                        SetText SPD, XmlSelectS.SPCNM(i), intRow, colSPECIMEN
                        
                        strWorkNo = XmlSelectS.WORKNO(i)
                        strWorkNo = Mid(strWorkNo, 1, 8) & "-" & Mid(strWorkNo, 9, 2) & "-" & Mid(strWorkNo, 11)
                        SetText SPD, strWorkNo, intRow, colCHARTNO
                        
                        strNames = GetText(SPD, intRow, colITEMS)
                        strNames = GetTestNm(XmlSelectS.TCLSCDLIST(i))
                        SetText SPD, strNames, intRow, colITEMS
    
'                        If gTest = "MTB" Then
'                            SetText SPD, "MTB", intRow, colPOSNO
'
'                        ElseIf gTest = "RP" Then
'                            SetText SPD, "RP", intRow, colPOSNO
'
'                        ElseIf gTest = "PB" Then
'                            SetText SPD, "PB", intRow, colPOSNO
'
'                        ElseIf gTest = "RPPB" Then
'                            If l = 0 Then
'                                SetText SPD, "RP", intRow, colPOSNO
'                            ElseIf l = 1 Then
'                                SetText SPD, "RP", intRow, colPOSNO
'                            Else
'                                SetText SPD, "PB", intRow, colPOSNO
'                            End If
'                        End If
    
                    End If
                End If
            End With
        Next
    Else
        MsgBox "��ȸ ����ڰ� �����ϴ�.", vbOKOnly + vbCritical, "��ũ����Ʈ ��ȸ"
    End If
    
'    If gTest = "RP" And l = 0 Then
'        l = l + 1
'        GoTo ReSearch
'    End If
'
'    If gTest = "RPPB" And l = 0 Then
'        l = l + 1
'        GoTo ReSearch
'    End If
'
'    If gTest = "RPPB" And l = 1 Then
'        l = l + 1
'        GoTo ReSearch
'    End If
    
    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMsg = "��    ġ : " & gHOSP.MACHNM & "_GetWorkList_NU" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0

End Sub


Public Sub GetWorkList_SUNGMO(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As Object)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim J           As Integer
    Dim k           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim sParam      As String
    Dim strTestCds  As String
    Dim sRcvData    As String
    Dim varRcvData  As Variant
    Dim varTstCode  As Variant
    Dim strNames    As String
    Dim strXmlName  As String
    Dim strWorkNo   As String
    
    Dim l As Integer
    
On Error GoTo RST
    
    Screen.MousePointer = 11
    SPD.MaxRows = 0
    
    blnSame = False
    strNames = ""
    l = 0
    
ReSearch:
    
    strTestCds = gHOSP.ORDCODE

    '��õ���𺴿�
    'http://emr016.cmcnu.or.kr/cmcnu/.live?submit_id=TRLII00101&business_id=li&instcd=016&eqmtcd=P08&startdd=20200514&enddd=20200514&

    sParam = ""
    sParam = sParam & "submit_id=TRLII00101&"                               'submit ID
    sParam = sParam & "business_id=li&"                                     'business_id
    sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"                        '����ڵ�
    sParam = sParam & "startdd=" & pFrom & "&"                              '�����۾�����
    sParam = sParam & "enddd=" & pTo & "&"                                  '�����۾�����
    sParam = sParam & "eqmtcd=" & gHOSP.MACHCD & "&"                        '����ڵ�
    
    sRcvData = OpenURLWithIE2(gHOSP.APIURL & sParam, frmMain.Inet1)
    
    Call SetSQLData("��ũ��ȸ", "Param:" & gHOSP.APIURL & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine)

    If InStr(1, sRcvData, "<?xml version") > 0 Then
        varRcvData = Split(sRcvData, "<worklist>")
    End If

    strXmlName = gHOSP.MACHNM & "_" & Format(CDate(Now), "yyyymmdd") & ".xml"

    Call SetXMLData(strXmlName, sRcvData)

    Call DisplayNode_InfoS(App.PATH & "\Xml\" & strXmlName, UBound(varRcvData))
    'Call DisplayNode_InfoS("D:\������Ʈ\VB\__NU\��õ���𺴿�_NU\����\worklist.xml", 10)

    Kill App.PATH & "\Xml\" & strXmlName
    
    If UBound(varRcvData) >= 1 Then
        For i = 0 To UBound(varRcvData) - 1
        'For i = 0 To 10
            With SPD
                .ReDraw = False
                blnSame = False
                
                '2019-12-11 �޸�
                '   rsltstat �� ó�� ���� �� ��...
                '   <rsltstat><![CDATA[-]]></rsltstat>
                '   <rsltstat><![CDATA[4]]></rsltstat>
                
                'If GetSampleTest_HDINFO(XmlSelectS.BCNO(i)) > 0 Then
                    
                If InStr(UCase(XmlSelectS.IFREQCDLIST(i)), "COVID19") > 0 Then
                    For J = 1 To SPD.DataRowCnt
                        strHospDate = GetText(SPD, J, colHOSPDATE)
                        strBarcode = GetText(SPD, J, colBARCODE)
                        If XmlSelectS.PRCPDD(i) & "" = strHospDate And XmlSelectS.BCNO(i) = strBarcode Then
                            blnSame = True
                            strNames = GetText(SPD, intRow, colITEMS)
                            strNames = strNames & "|" & GetTestNm(XmlSelectS.TestCd(i))
                            
                            SetText SPD, strNames, intRow, colITEMS
                            strNames = ""
                        End If
                    Next
                    
                    If blnSame = False Then
                        .MaxRows = .MaxRows + 1
                        intRow = .MaxRows
                
                        SetText SPD, "1", intRow, colCHECKBOX
                        SetText SPD, CStr(intRow), intRow, colSEQNO
                        SetText SPD, XmlSelectS.SPCACPTDT(i), intRow, colHOSPDATE
                        SetText SPD, XmlSelectS.BCNO(i), intRow, colBARCODE
                        SetText SPD, XmlSelectS.PID(i), intRow, colPID
                        SetText SPD, XmlSelectS.PATNM(i), intRow, colPNAME
                        SetText SPD, XmlSelectS.SEX(i), intRow, colPSEX
                        SetText SPD, XmlSelectS.AGE(i), intRow, colPAGE
                        SetText SPD, XmlSelectS.SPCNM(i), intRow, colSPECIMEN
                        
                        strWorkNo = XmlSelectS.WORKNO(i)
                        strWorkNo = Mid(strWorkNo, 1, 8) & "-" & Mid(strWorkNo, 9, 2) & "-" & Mid(strWorkNo, 11)
                        SetText SPD, strWorkNo, intRow, colCHARTNO
                        
                        strNames = GetText(SPD, intRow, colITEMS)
                        strNames = GetTestNm(XmlSelectS.TCLSCDLIST(i))
                        SetText SPD, strNames, intRow, colITEMS
    
'                        If gTest = "MTB" Then
'                            SetText SPD, "MTB", intRow, colPOSNO
'
'                        ElseIf gTest = "RP" Then
'                            SetText SPD, "RP", intRow, colPOSNO
'
'                        ElseIf gTest = "PB" Then
'                            SetText SPD, "PB", intRow, colPOSNO
'
'                        ElseIf gTest = "RPPB" Then
'                            If l = 0 Then
'                                SetText SPD, "RP", intRow, colPOSNO
'                            ElseIf l = 1 Then
'                                SetText SPD, "RP", intRow, colPOSNO
'                            Else
'                                SetText SPD, "PB", intRow, colPOSNO
'                            End If
'                        End If
    
                    End If
                End If
            End With
        Next
    Else
        MsgBox "��ȸ ����ڰ� �����ϴ�.", vbOKOnly + vbCritical, "��ũ����Ʈ ��ȸ"
    End If
    
'    If gTest = "RP" And l = 0 Then
'        l = l + 1
'        GoTo ReSearch
'    End If
'
'    If gTest = "RPPB" And l = 0 Then
'        l = l + 1
'        GoTo ReSearch
'    End If
'
'    If gTest = "RPPB" And l = 1 Then
'        l = l + 1
'        GoTo ReSearch
'    End If
    
    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMsg = "��    ġ : " & gHOSP.MACHNM & "_GetWorkList_NU" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0

End Sub

Public Sub GetWorkList_KHNMC(ByVal pFrom As String, ByVal pTo As String, ByVal SPD As Object)
    Dim RS          As ADODB.Recordset
    Dim blnSame     As Boolean
    
    Dim i           As Integer
    Dim J           As Integer
    Dim k           As Integer
    Dim iCnt        As Integer
    Dim intRow      As Integer
    Dim strHospDate As String
    Dim strBarcode  As String
    Dim sParam      As String
    Dim strTestCds  As String
    Dim sRcvData    As String
    Dim varRcvData  As Variant
    Dim varRcvDtl   As Variant
    Dim varTstCode  As Variant
    Dim strNames    As String
    Dim strXmlName  As String
    Dim strWorkNo   As String
    Dim l As Integer
    Dim pDate       As String
    
    'Dim strXML(4)   As String
    Dim strPID      As String
    Dim strKORNM    As String
    Dim strGSNO     As String
    Dim strBCNO     As String
    Dim strIFCODE   As String
    Dim strRSTFLAG  As String
    
    Dim blnAdd      As Boolean
    Dim strAtNm     As String
    Dim sRecData    As String

On Error GoTo RST
    
    Screen.MousePointer = 11
    SPD.MaxRows = 0
    
    blnSame = False
    strNames = ""
    l = 0
    
ReSearch:
    
'http://testdev.khnmc.or.kr/lis/jindangeomsaweb/GyeolGwaIF.live?Mode=reqGetIFPCR&Data1=R04&Data2=20200810
'Data1: ����ڵ�
'Data2: ��������
    
    If pFrom = pTo Then
        pDate = pFrom
    Else
    
    End If
    
    'strTestCds = gHOSP.ORDCODE

    sParam = ""
    sParam = sParam & "Data1=" & gHOSP.MACHCD & "&"     '����ڵ�
    sParam = sParam & "Data2=" & pFrom & "&"            '��������
    
    sRcvData = OpenURLWithIE2("http://" & gHOSP.APIURL & "/lis/jindangeomsaweb/GyeolGwaIF.live?Mode=reqGetIFPCR&" & sParam, frmMain.Inet1)
    
    Call SetSQLData("��ũ��ȸ", "Return==>" & sRcvData & vbNewLine, "A")



'sRecData = ""
'sRecData = sRecData & "<?xml version=""1.0"" encoding=""euc-kr"" ?>" & vbCrLf
'sRecData = sRecData & "" & vbCrLf
'sRecData = sRecData & "" & vbCrLf
'sRecData = sRecData & "" & vbCrLf
'sRecData = sRecData & "" & vbCrLf
'sRecData = sRecData & "" & vbCrLf
'sRecData = sRecData & "" & vbCrLf
'sRecData = sRecData & "" & vbCrLf
'sRecData = sRecData & "" & vbCrLf
'sRecData = sRecData & "" & vbCrLf
'sRecData = sRecData & "" & vbCrLf
'sRecData = sRecData & "<root>" & vbCrLf
'sRecData = sRecData & "    " & vbCrLf
'sRecData = sRecData & "" & vbCrLf
'sRecData = sRecData & "" & vbCrLf
'sRecData = sRecData & "   " & vbCrLf
'sRecData = sRecData & "       " & vbCrLf
'sRecData = sRecData & "    " & vbCrLf
'sRecData = sRecData & "" & vbCrLf
'sRecData = sRecData & "   " & vbCrLf
'sRecData = sRecData & "        " & vbCrLf
'sRecData = sRecData & "        " & vbCrLf
'sRecData = sRecData & "           <List>" & vbCrLf
'sRecData = sRecData & "               <pid>08686896</pid>" & vbCrLf
'sRecData = sRecData & "               <kornm>������B</kornm>" & vbCrLf
'sRecData = sRecData & "               <gsno>RV16-20-00001</gsno>" & vbCrLf
'sRecData = sRecData & "               <bcno>253430060</bcno>              " & vbCrLf
'sRecData = sRecData & "               <ifcode>Adeno</ifcode>              " & vbCrLf
'sRecData = sRecData & "           </List>" & vbCrLf
'sRecData = sRecData & "        " & vbCrLf
'sRecData = sRecData & "           <List>" & vbCrLf
'sRecData = sRecData & "               <pid>08686896</pid>" & vbCrLf
'sRecData = sRecData & "               <kornm>������B</kornm>" & vbCrLf
'sRecData = sRecData & "               <gsno>RV16-20-00001</gsno>" & vbCrLf
'sRecData = sRecData & "               <bcno>253430060</bcno>              " & vbCrLf
'sRecData = sRecData & "               <ifcode>InfluA</ifcode>             " & vbCrLf
'sRecData = sRecData & "           </List>" & vbCrLf
'sRecData = sRecData & "        " & vbCrLf
'sRecData = sRecData & "           <List>" & vbCrLf
'sRecData = sRecData & "               <pid>08686896</pid>" & vbCrLf
'sRecData = sRecData & "               <kornm>������B</kornm>" & vbCrLf
'sRecData = sRecData & "               <gsno>RV16-20-00001</gsno>" & vbCrLf
'sRecData = sRecData & "               <bcno>253430060</bcno>              " & vbCrLf
'sRecData = sRecData & "               <ifcode>Parainflu1</ifcode>             " & vbCrLf
'sRecData = sRecData & "           </List>" & vbCrLf
'sRecData = sRecData & "        " & vbCrLf
'sRecData = sRecData & "           <List>" & vbCrLf
'sRecData = sRecData & "               <pid>08686896</pid>" & vbCrLf
'sRecData = sRecData & "               <kornm>������B</kornm>" & vbCrLf
'sRecData = sRecData & "               <gsno>RV16-20-00001</gsno>"
'sRecData = sRecData & "               <bcno>253430060</bcno>              "
'sRecData = sRecData & "               <ifcode>Parainflu2</ifcode>             "
'sRecData = sRecData & "           </List>"
'sRecData = sRecData & "        "
'sRecData = sRecData & "           <List>"
'sRecData = sRecData & "               <pid>08686896</pid>"
'sRecData = sRecData & "               <kornm>������B</kornm>"
'sRecData = sRecData & "               <gsno>RV16-20-00001</gsno>"
'sRecData = sRecData & "               <bcno>253430060</bcno>              "
'sRecData = sRecData & "               <ifcode>Parainflu3</ifcode>             "
'sRecData = sRecData & "           </List>"
'sRecData = sRecData & "        "
'sRecData = sRecData & "           <List>"
'sRecData = sRecData & "               <pid>08686896</pid>"
'sRecData = sRecData & "               <kornm>������B</kornm>"
'sRecData = sRecData & "               <gsno>RV16-20-00001</gsno>"
'sRecData = sRecData & "               <bcno>253430060</bcno>              "
'sRecData = sRecData & "               <ifcode>Parainflu4</ifcode>             "
'sRecData = sRecData & "           </List>"
'sRecData = sRecData & "        "
'sRecData = sRecData & "           <List>"
'sRecData = sRecData & "               <pid>08686896</pid>"
'sRecData = sRecData & "               <kornm>������B</kornm>"
'sRecData = sRecData & "               <gsno>RV16-20-00001</gsno>"
'sRecData = sRecData & "               <bcno>253430060</bcno>              "
'sRecData = sRecData & "               <ifcode>Rhino</ifcode>              "
'sRecData = sRecData & "           </List>"
'sRecData = sRecData & "        "
'sRecData = sRecData & "           <List>"
'sRecData = sRecData & "               <pid>08686896</pid>"
'sRecData = sRecData & "               <kornm>������B</kornm>"
'sRecData = sRecData & "               <gsno>RV16-20-00001</gsno>"
'sRecData = sRecData & "               <bcno>253430060</bcno>              "
'sRecData = sRecData & "               <ifcode>RSVA</ifcode>               "
'sRecData = sRecData & "           </List>"
'sRecData = sRecData & "        "
'sRecData = sRecData & "           <List>"
'sRecData = sRecData & "               <pid>08686896</pid>"
'sRecData = sRecData & "               <kornm>������B</kornm>"
'sRecData = sRecData & "               <gsno>RV16-20-00001</gsno>"
'sRecData = sRecData & "               <bcno>253430060</bcno>              "
'sRecData = sRecData & "               <ifcode>RSVB</ifcode>               "
'sRecData = sRecData & "           </List>"
'sRecData = sRecData & "        "
'sRecData = sRecData & "           <List>"
'sRecData = sRecData & "               <pid>08686896</pid>"
'sRecData = sRecData & "               <kornm>������B</kornm>"
'sRecData = sRecData & "               <gsno>RV16-20-00001</gsno>"
'sRecData = sRecData & "               <bcno>253430060</bcno>              "
'sRecData = sRecData & "               <ifcode>Boca</ifcode>               "
'sRecData = sRecData & "           </List>"
'sRecData = sRecData & "        "
'sRecData = sRecData & "           <List>"
'sRecData = sRecData & "               <pid>08686896</pid>"
'sRecData = sRecData & "               <kornm>������B</kornm>"
'sRecData = sRecData & "               <gsno>RV16-20-00001</gsno>"
'sRecData = sRecData & "               <bcno>253430060</bcno>              "
'sRecData = sRecData & "               <ifcode>Metapneumo</ifcode>             "
'sRecData = sRecData & "           </List>"
'sRecData = sRecData & "        "
'sRecData = sRecData & "           <List>"
'sRecData = sRecData & "               <pid>08686896</pid>"
'sRecData = sRecData & "               <kornm>������B</kornm>"
'sRecData = sRecData & "               <gsno>RV16-20-00001</gsno>"
'sRecData = sRecData & "               <bcno>253430060</bcno>              "
'sRecData = sRecData & "               <ifcode>Corona229E</ifcode>             "
'sRecData = sRecData & "           </List>"
'sRecData = sRecData & "        "
'sRecData = sRecData & "           <List>"
'sRecData = sRecData & "               <pid>08686896</pid>"
'sRecData = sRecData & "               <kornm>������B</kornm>"
'sRecData = sRecData & "               <gsno>RV16-20-00001</gsno>"
'sRecData = sRecData & "               <bcno>253430060</bcno>              "
'sRecData = sRecData & "               <ifcode>CoronaNL63</ifcode>             "
'sRecData = sRecData & "           </List>"
'sRecData = sRecData & "        "
'sRecData = sRecData & "           <List>"
'sRecData = sRecData & "               <pid>08686896</pid>"
'sRecData = sRecData & "               <kornm>������B</kornm>"
'sRecData = sRecData & "               <gsno>RV16-20-00001</gsno>"
'sRecData = sRecData & "               <bcno>253430060</bcno>              "
'sRecData = sRecData & "               <ifcode>CoronaOC43</ifcode>             "
'sRecData = sRecData & "           </List>"
'sRecData = sRecData & "        "
'sRecData = sRecData & "           <List>"
'sRecData = sRecData & "               <pid>08686896</pid>"
'sRecData = sRecData & "               <kornm>������B</kornm>"
'sRecData = sRecData & "               <gsno>RV16-20-00001</gsno>"
'sRecData = sRecData & "               <bcno>253430060</bcno>              "
'sRecData = sRecData & "               <ifcode>Entero</ifcode>             "
'sRecData = sRecData & "           </List>"
'sRecData = sRecData & "        "
'sRecData = sRecData & "           <List>"
'sRecData = sRecData & "               <pid>08686896</pid>"
'sRecData = sRecData & "               <kornm>������B</kornm>"
'sRecData = sRecData & "               <gsno>RV16-20-00001</gsno>"
'sRecData = sRecData & "               <bcno>253430060</bcno>              "
'sRecData = sRecData & "               <ifcode>InfluB</ifcode>             "
'sRecData = sRecData & "           </List>"
'sRecData = sRecData & "        "
'sRecData = sRecData & "           <List>"
'sRecData = sRecData & "               <pid>05231302</pid>"
'sRecData = sRecData & "               <kornm>���õ</kornm>"
'sRecData = sRecData & "               <gsno>����corona-20-14397</gsno>"
'sRecData = sRecData & "               <bcno>153430010</bcno>              "
'sRecData = sRecData & "               <ifcode>COVID19</ifcode>                "
'sRecData = sRecData & "           </List>"
'sRecData = sRecData & "        "
'sRecData = sRecData & "           <List>"
'sRecData = sRecData & "               <pid>02579295</pid>"
'sRecData = sRecData & "               <kornm>ǥ����</kornm>"
'sRecData = sRecData & "               <gsno>����corona-20-14398</gsno>"
'sRecData = sRecData & "               <bcno>153430020</bcno>              "
'sRecData = sRecData & "               <ifcode>COVID19</ifcode>                "
'sRecData = sRecData & "           </List>"
'sRecData = sRecData & "        "
'sRecData = sRecData & "           <List>"
'sRecData = sRecData & "               <pid>08544979</pid>"
'sRecData = sRecData & "               <kornm>������</kornm>"
'sRecData = sRecData & "               <gsno>����corona-20-14399</gsno>"
'sRecData = sRecData & "               <bcno>153430030</bcno>              "
'sRecData = sRecData & "               <ifcode>COVID19</ifcode>                "
'sRecData = sRecData & "           </List>"
'sRecData = sRecData & "        "
'sRecData = sRecData & "           <List>"
'sRecData = sRecData & "               <pid>08667942</pid>"
'sRecData = sRecData & "               <kornm>������</kornm>"
'sRecData = sRecData & "               <gsno>����corona-20-14400</gsno>"
'sRecData = sRecData & "               <bcno>253430010</bcno>              "
'sRecData = sRecData & "               <ifcode>COVID19</ifcode>                "
'sRecData = sRecData & "           </List>"
'sRecData = sRecData & "        "
'sRecData = sRecData & "           <List>"
'sRecData = sRecData & "               <pid>08686275</pid>"
'sRecData = sRecData & "               <kornm>������</kornm>"
'sRecData = sRecData & "               <gsno>����corona-20-14401</gsno>"
'sRecData = sRecData & "               <bcno>253430020</bcno>              "
'sRecData = sRecData & "               <ifcode>COVID19</ifcode>                "
'sRecData = sRecData & "           </List>"
'sRecData = sRecData & "        "
'sRecData = sRecData & "           <List>"
'sRecData = sRecData & "               <pid>08234929</pid>"
'sRecData = sRecData & "               <kornm>������</kornm>"
'sRecData = sRecData & "               <gsno>����corona-20-14402</gsno>"
'sRecData = sRecData & "               <bcno>253430030</bcno>              "
'sRecData = sRecData & "               <ifcode>COVID19</ifcode>                "
'sRecData = sRecData & "           </List>"
'sRecData = sRecData & "        "
'sRecData = sRecData & "           <List>"
'sRecData = sRecData & "               <pid>07614961</pid>"
'sRecData = sRecData & "               <kornm>������</kornm>"
'sRecData = sRecData & "               <gsno>����corona-20-14403</gsno>"
'sRecData = sRecData & "               <bcno>253430040</bcno>              "
'sRecData = sRecData & "               <ifcode>COVID19</ifcode>                "
'sRecData = sRecData & "           </List>"
'sRecData = sRecData & "        "
'sRecData = sRecData & "           <List>"
'sRecData = sRecData & "               <pid>08651067</pid>"
'sRecData = sRecData & "               <kornm>�����</kornm>"
'sRecData = sRecData & "               <gsno>����corona-20-14404</gsno>"
'sRecData = sRecData & "               <bcno>253430050</bcno>              "
'sRecData = sRecData & "               <ifcode>COVID19</ifcode>                "
'sRecData = sRecData & "           </List>"
'sRecData = sRecData & "        "
'sRecData = sRecData & "           <List>"
'sRecData = sRecData & "               <pid>01111001</pid>"
'sRecData = sRecData & "               <kornm>���ͱ�</kornm>"
'sRecData = sRecData & "               <gsno>����corona-20-14405</gsno>"
'sRecData = sRecData & "               <bcno>153430040</bcno>              "
'sRecData = sRecData & "               <ifcode>COVID19</ifcode>                "
'sRecData = sRecData & "           </List>"
'sRecData = sRecData & "        "
'sRecData = sRecData & "    "
'sRecData = sRecData & "</root>" & vbCrLf
'
'sRcvData = sRecData

    If InStr(1, sRcvData, "<?xml version") > 0 Then
        varRcvData = Split(sRcvData, "<List>")
    End If


    'strXmlName = gHOSP.MACHNM & "_" & Format(CDate(Now), "yyyymmdd") & ".xml"
    'Call SetXMLData(strXmlName, sRcvData)
    'Call DisplayNode_InfoS_Worklist_KHNMC(App.PATH & "\Xml\" & strXmlName, UBound(varRcvData))
    'Call DisplayNode_InfoS_Worklist_KHNMC("D:\������Ʈ\VB\��������뺴��\IF\XML\worklist.xml", 10)

'    Kill App.PATH & "\Xml\" & strXmlName
    
    If UBound(varRcvData) >= 1 Then
        For i = 1 To UBound(varRcvData)
            varRcvData(i) = Replace(varRcvData(i), vbLf, "")
            varRcvDtl = Split(varRcvData(i), vbCr)
            
            For J = 1 To 6
                strAtNm = Trim(mGetP(varRcvDtl(J), 1, ">"))
                strAtNm = UCase(Mid(strAtNm, 2))
                Select Case strAtNm
                    Case "PID":     strPID = Trim(mGetP(mGetP(varRcvDtl(J), 2, ">"), 1, "<"))
                    Case "KORNM":   strKORNM = Trim(mGetP(mGetP(varRcvDtl(J), 2, ">"), 1, "<"))
                    Case "GSNO":    strGSNO = Trim(mGetP(mGetP(varRcvDtl(J), 2, ">"), 1, "<"))
                    Case "BCNO":    strBCNO = Trim(mGetP(mGetP(varRcvDtl(J), 2, ">"), 1, "<"))
                    Case "IFCODE":  strPID = Trim(mGetP(mGetP(varRcvDtl(J), 2, ">"), 1, "<"))
                    Case "RSTFLAG": strRSTFLAG = Trim(mGetP(mGetP(varRcvDtl(J), 2, ">"), 1, "<"))
                End Select
            Next
            
'            strPID = Trim(mGetP(mGetP(varRcvDtl(1), 2, ">"), 1, "<"))
'            strKORNM = Trim(mGetP(mGetP(varRcvDtl(2), 2, ">"), 1, "<"))
'            strGSNO = Trim(mGetP(mGetP(varRcvDtl(3), 2, ">"), 1, "<"))
'            strBCNO = Trim(mGetP(mGetP(varRcvDtl(4), 2, ">"), 1, "<"))
'            strIFCODE = Trim(mGetP(mGetP(varRcvDtl(5), 2, ">"), 1, "<"))
    
            blnAdd = False
            
            
            If frmMain.optTest(0).Value = True Then
                'covid19
                If InStr(strGSNO, "corona") > 0 Then
                    blnAdd = True
                End If
                If InStr(strGSNO, "COVID-19") > 0 Then
                    blnAdd = True
                End If
            ElseIf frmMain.optTest(1).Value = True Then
                'rv16
                If InStr(strGSNO, "RV16") > 0 Then
                    blnAdd = True
                End If
            ElseIf frmMain.optTest(2).Value = True Then
                'pb5
                If InStr(strGSNO, "Pneumo") > 0 Then
                    blnAdd = True
                End If
            End If
            
            varRcvData(i) = Replace(varRcvData(i), vbLf, "")
            varRcvDtl = Split(varRcvData(i), vbCr)
            
            If blnAdd = True Then
                With SPD
                    .ReDraw = False
                    blnSame = False
    
                    For J = 1 To SPD.DataRowCnt
                        strHospDate = GetText(SPD, J, colHOSPDATE)
                        strBarcode = GetText(SPD, J, colBARCODE)
                        If strBCNO = strBarcode Then
                            blnSame = True
                            strNames = GetText(SPD, intRow, colITEMS)
                            strNames = strNames & "|" & GetTestNm(strIFCODE)
    
                            SetText SPD, strNames, intRow, colITEMS
                            strNames = ""
                        End If
                    Next
    
                    If blnSame = False Then
                        .MaxRows = .MaxRows + 1
                        intRow = .MaxRows
    
                        SetText SPD, "1", intRow, colCHECKBOX
                        SetText SPD, CStr(intRow), intRow, colSEQNO
                        SetText SPD, pFrom, intRow, colHOSPDATE
                        SetText SPD, strBCNO, intRow, colBARCODE
                        SetText SPD, strPID, intRow, colPID
                        SetText SPD, strKORNM, intRow, colPNAME
                        'SetText SPD, strGSNO, intRow, colCHARTNO
                        
                        If InStr(strGSNO, "COVID-19") > 0 Then
                            SetText SPD, Mid(strGSNO, InStr(strGSNO, "COVID-19-") + 9), intRow, colCHARTNO
                        Else
                            SetText SPD, Mid(strGSNO, InStr(strGSNO, "-") + 1), intRow, colCHARTNO
                        End If
                        strNames = GetText(SPD, intRow, colITEMS)
                        strNames = GetTestNm_ORD(strIFCODE)
                        
                        SetText SPD, strNames, intRow, colITEMS
    
                        If strRSTFLAG = "1" Then
                            .Row = intRow
                            .Col = colPNAME
                            .BackColor = vbYellow
                        ElseIf strRSTFLAG = "2" Then
                            .Row = intRow
                            .Col = colPNAME
                            .BackColor = vbYellow
                        Else
                            .Row = intRow
                            .Col = colPNAME
                            .BackColor = vbWhite
                        End If
    
                    End If
                End With
            End If
        Next
    Else
        MsgBox "��ȸ ����ڰ� �����ϴ�.", vbOKOnly + vbCritical, "��ũ����Ʈ ��ȸ"
    End If
    
'    If gTest = "RP" And l = 0 Then
'        l = l + 1
'        GoTo ReSearch
'    End If
'
'    If gTest = "RPPB" And l = 0 Then
'        l = l + 1
'        GoTo ReSearch
'    End If
'
'    If gTest = "RPPB" And l = 1 Then
'        l = l + 1
'        GoTo ReSearch
'    End If
    
    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True
    
    Screen.MousePointer = 0

Exit Sub

RST:
     
                strErrMsg = "��    ġ : " & gHOSP.MACHNM & "_GetWorkList_KHNMC" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show 'vbModal
    
    Screen.MousePointer = 0

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

'    SQL = SQL & "     , c.�������                          AS STATE        " & vbCrLf

    SQL = ""
    SQL = SQL & "Select DISTINCT "
    SQL = SQL & "       (a.����� + a.����� + a.������)    AS HOSPDATE     " & vbCrLf
    SQL = SQL & "     , a.íƮ��ȣ                          AS CHARTNO      " & vbCrLf
    SQL = SQL & "     , b.�����ڸ�                          AS PNAME        " & vbCrLf
    SQL = SQL & "     , b.�ֹε�Ϲ�ȣ                      AS PJUMIN       " & vbCrLf
    SQL = SQL & "     , COUNT(a.ó���ڵ�)                   AS CNT          " & vbCrLf
    SQL = SQL & "  From TB_�˻��׸� a                                       " & vbCrLf
    SQL = SQL & "     , TB_�������� b                                       " & vbCrLf
    SQL = SQL & "     , TB_����⺻ c                                       " & vbCrLf
    SQL = SQL & " Where (a.����� + a.����� + a.������) >= '" & pFrom & "' " & vbCrLf
    SQL = SQL & "   And (a.����� + a.����� + a.������) <= '" & pTo & "'   " & vbCrLf
    SQL = SQL & "   And a.ó���ȣ > 0                                      " & vbCrLf
    SQL = SQL & "   And c.������� IN ('1','5','6','7','8','9')             " & vbCrLf
    'SQL = SQL & "   And (a.ó���ڵ� + a.�����ڵ�) IN (" & gAllTestCd & ")   " & vbCrLf
    SQL = SQL & "   And (a.ó���ڵ� + '|' + a.�����ڵ�) IN (" & gAllTestCd & ")   " & vbCrLf
    SQL = SQL & "   And (a.�˻��� IS NULL OR a.�˻��� = '')             " & vbCrLf
    SQL = SQL & "   And a.�����    = c.�����                              " & vbCrLf
    SQL = SQL & "   And a.�����    = c.�����                              " & vbCrLf
    SQL = SQL & "   And a.������    = c.������                              " & vbCrLf
    SQL = SQL & "   And a.íƮ��ȣ  = c.íƮ��ȣ                            " & vbCrLf
    SQL = SQL & "   And a.íƮ��ȣ  = b.íƮ��ȣ                            " & vbCrLf
    SQL = SQL & "   And (a.�˻��� IS NULL OR a.�˻��� = '')             " & vbCrLf
    SQL = SQL & " GROUP BY HOSPDATE, a.íƮ��ȣ, b.�����ڸ�, b.�ֹε�Ϲ�ȣ " & vbCrLf   ', c.�������
    SQL = SQL & " Order By a.�����, a.�����, a.������, b.�����ڸ�         " & vbCrLf

    Call SetSQLData("��ũ��ȸ", SQL, "")

    '-- Record Count ������
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
        frmMain.lblComStatus.Caption = "��ũ����Ʈ ��ȸ ����ڰ� �����ϴ�."
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_GetWorkList_MEDICHART" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
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
    SQL = SQL & "       J.��������          AS HOSPDATE                         " & vbCrLf
    SQL = SQL & "     , L.��ü��ȣ          AS BARCODE                          " & vbCrLf
    SQL = SQL & "     , A.íƮ��ȣ          AS CHARTNO                          " & vbCrLf
    SQL = SQL & "     , J.������ȣ          AS PID                              " & vbCrLf
    SQL = SQL & "     , A.ȯ���̸�          AS PNAME                            " & vbCrLf
    SQL = SQL & "     , A.ȯ�ڼ���          AS SEX                              " & vbCrLf
    SQL = SQL & "     , A.ȯ�ڳ���          AS AGE                              " & vbCrLf
    SQL = SQL & "     , COUNT(L.ó���ڵ�)   AS CNT                              " & vbCrLf
    SQL = SQL & "  FROM         TB_����˻� L                                   " & vbCrLf
    SQL = SQL & "   INNER JOIN  TB_�������� J ON (L.��������ID = J.��������ID)  " & vbCrLf
    SQL = SQL & "   INNER JOIN  TB_�����Ϲ� A ON (J.��������   = A.��������     " & vbCrLf
    SQL = SQL & "                            AND  J.íƮ��ȣ   = A.íƮ��ȣ     " & vbCrLf
    SQL = SQL & "                            AND  J.�����ȣ   = A.�����ȣ)    " & vbCrLf
    SQL = SQL & " Where J.�������� BETWEEN '" & pFrom & "' and '" & pTo & "'    " & vbCrLf
    SQL = SQL & "   AND L.�˻���� < 5                                          " & vbCrLf
    SQL = SQL & "   AND L.ó���ڵ� + L.�����ڵ� IN (" & gAllTestCd & ")         " & vbCrLf
    SQL = SQL & " GROUP BY J.��������, L.��ü��ȣ, A.íƮ��ȣ, J.������ȣ, A.ȯ���̸�, A.ȯ�ڼ���, A.ȯ�ڳ��� " & vbCrLf
'    SQL = SQL & " ORDER BY J.��������, L.��ü��ȣ                               " & vbCrLf

    Call SetSQLData("��ũ��ȸ", SQL, "")

    '-- Record Count ������
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
                    
                    '��񿡼� ������û�� �ȿ��� ��ġ������
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
        frmMain.lblComStatus.Caption = "��ũ����Ʈ ��ȸ ����ڰ� �����ϴ�."
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_GetWorkList_KCHART" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    
    frmErrMsg.Show vbModal

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
    SQL = SQL & "       CONVERT(NVARCHAR(50),M.��������,112)    AS HOSPDATE                                 " & vbCrLf
    SQL = SQL & "     , M.������ȣ                              AS PID                                      " & vbCrLf
    SQL = SQL & "     , M.��Ʈ��ȣ                              AS CHARTNO                                  " & vbCrLf
    SQL = SQL & "     , M.����                                  AS PNAME                                    " & vbCrLf
    SQL = SQL & "     , M.����                                  AS SEX                                      " & vbCrLf
    SQL = SQL & "     , M.����                                  AS AGE                                      " & vbCrLf
    SQL = SQL & "     , M.�ŷ�ó��                              AS DEPT                                     " & vbCrLf
    SQL = SQL & "     , E.�˻��ڵ�                              AS ITEM                                     " & vbCrLf
    'SQL = SQL & "     , COUNT(E.�˻��ڵ�)                       AS CNT                                      " & vbCrLf
    SQL = SQL & "  FROM VW_�˻����� M                                                                       " & vbCrLf
    SQL = SQL & "     , VW_�˻��� R                                                                       " & vbCrLf
    SQL = SQL & "     , VW_�˻��ڵ� E                                                                       " & vbCrLf
    SQL = SQL & " WHERE M.�������� BETWEEN CONVERT(DATE, '" & pFrom & "') AND CONVERT(DATE, '" & pTo & "')  " & vbCrLf
    SQL = SQL & "   AND M.��������      = R.��������                                                        " & vbCrLf
    SQL = SQL & "   AND M.������ȣ      = R.������ȣ                                                        " & vbCrLf
    SQL = SQL & "   AND R.�˻��ڵ�      = E.�˻��ڵ�                                                        " & vbCrLf
    SQL = SQL & "   AND E.�к��ڵ�      = '" & gHOSP.PARTCD & "'                                            " & vbCrLf    'U2
    SQL = SQL & "   AND E.�˻��ڵ�      IN (" & gAllTestCd & ")                                             " & vbCrLf
    SQL = SQL & "   AND ISNULL(R.������, 'N') <> 'Y'                                                      " & vbCrLf
    SQL = SQL & "   AND (R.����� IS NULL OR R.����� = '')                                                 " & vbCrLf
    'SQL = SQL & " GROUP BY M.��������,M.������ȣ,M.��Ʈ��ȣ,M.����,M.����,M.����,M.�ŷ�ó��                 " & vbCrLf
    
    Call SetSQLData("��ũ��ȸ", SQL, "")

    '-- Record Count ������
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
                    
                    '��񿡼� ������û�� �ȿ��� ��ġ������
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
        frmMain.lblComStatus.Caption = "��ũ����Ʈ ��ȸ ����ڰ� �����ϴ�."
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_GetWorkList_AMIS" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
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
    
'    SQL = SQL & "     , L.LABINSNUM     as      ó�����    " & vbCrLf
'    SQL = SQL & "     , L.LABSMPNAM     as      ��ü��      " & vbCrLf
    
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
    SQL = SQL & "   AND L.LABKEYNUM     = D.DATKEYNUM                           " & vbCrLf      '-- ���̺���Ű��
    SQL = SQL & "   AND L.LABATTEND     = D.DATATTEND                           " & vbCrLf      '-- ������ȣ
    SQL = SQL & "   AND L.LABATTEND     = M.MANATTEND                           " & vbCrLf      '-- ������ȣ
    SQL = SQL & "   AND L.LABCHTNUM     = D.DATCHTNUM                           " & vbCrLf      '-- íƮ��ȣ
    SQL = SQL & "   AND L.LABCHTNUM     = M.MANCHTNUM                           " & vbCrLf      '-- íƮ��ȣ
    SQL = SQL & "   AND L.LABODRDTE     = D.DATODRDTE                           " & vbCrLf      '-- ó������
    SQL = SQL & "   AND L.LABODRCOD     IN (" & gAllTestCd & ")                 " & vbCrLf      '-- ó��˻��׸�
    SQL = SQL & "   AND (L.LABCANCEL    = '' OR L.LABCANCEL IS NULL)            " & vbCrLf      '-- ��ҿ���
    SQL = SQL & "   AND (L.LABRESULT    = '' OR L.LABRESULT IS NULL)            " & vbCrLf      '-- �˻���
    SQL = SQL & "   AND L.LABENDDEP     < '3'                                   " & vbCrLf      '-- ó������ (2:����, 3:����Է�)
'    SQL = SQL & " ORDER BY L.LABODRDTE, L.LABCHTNUM, L.LABBARCOD, L.LABINSNUM   " & vbCrLf

    Call SetSQLData("��ũ��ȸ", SQL, "")

    '-- Record Count ������
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
                        Case "A":   SetText SPD, "�ܷ�", intRow, colINOUT
                        Case "F":   SetText SPD, "�Կ�", intRow, colINOUT
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
        frmMain.lblComStatus.Caption = "��ũ����Ʈ ��ȸ ����ڰ� �����ϴ�."
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

    Screen.MousePointer = 0

Exit Sub

ErrHandle:
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_GetWorkList_BIT70" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
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
    
    Call SetSQLData("��ũ��ȸ", SQL, "")

    '-- Record Count ������
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
'                        SetText SPD, frmMain.txtSeqNo.Text, intRow, colSEQNO
'                        frmMain.txtSeqNo.Text = frmMain.txtSeqNo.Text + 1
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
'            frmWorkList.lblStatus.Caption = ">> ��ȸ ����ڰ� �����ϴ�."
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "Form_Load" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show vbModal

End Sub


'-- �������� Ű ��������
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
                SQL = SQL & " WHERE SPECIMENNUM = '" & pBarcode & "'  " & vbCrLf
'                SQL = SQL & "   AND LABCODE     = '" & pTestCd & "'     " & vbCrLf
                
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
                SQL = SQL & "       L.����˻�ID                AS ORDCODE                      " & vbCrLf
                SQL = SQL & "     , L.��������ID                AS TESTSUBCODE                  " & vbCrLf
                SQL = SQL & "  FROM             TB_����˻� L                                   " & vbCrLf
                SQL = SQL & "       INNER JOIN  TB_�������� J ON  (L.��������ID = J.��������ID) " & vbCrLf
                SQL = SQL & "       INNER JOIN  TB_�����Ϲ� A ON  (J.��������   = A.��������    " & vbCrLf
                SQL = SQL & "                                AND   J.íƮ��ȣ   = A.íƮ��ȣ    " & vbCrLf
                SQL = SQL & "                                AND   J.�����ȣ   = A.�����ȣ)   " & vbCrLf
                SQL = SQL & " Where L.��ü��ȣ= '" & pBarcode & "'                              " & vbCrLf
                SQL = SQL & "   AND L.�˻���� < 5                                              " & vbCrLf
                SQL = SQL & "   AND (L.ó���ڵ� + L.�����ڵ�) = '" & pTestCd & "'               " & vbCrLf
    
                Call SetSQLData("SUB�ڵ���ȸ", SQL)
                
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
                SQL = SQL & "   AND O.CLAS          = 4                         " & vbCrLf '�ӻ󺴸�
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
                SQL = SQL & "   AND L.LABKEYNUM  = D.DATKEYNUM                  " & vbCrLf    '-- ���̺���Ű��
                SQL = SQL & "   AND L.LABATTEND  = D.DATATTEND                  " & vbCrLf    '-- ������ȣ
                SQL = SQL & "   AND L.LABATTEND  = M.MANATTEND                  " & vbCrLf    '-- ������ȣ ???
                SQL = SQL & "   AND L.LABCHTNUM  = D.DATCHTNUM                  " & vbCrLf    '-- íƮ��ȣ
                SQL = SQL & "   AND L.LABCHTNUM  = M.MANCHTNUM                  " & vbCrLf    '-- íƮ��ȣ
                SQL = SQL & "   AND L.LABODRDTE  = D.DATODRDTE                  " & vbCrLf    '-- ó������
                SQL = SQL & "   AND L.LABODRCOD IN (" & gAllTestCd & ")         " & vbCrLf
                SQL = SQL & "   AND (L.LABCANCEL = '' OR L.LABCANCEL IS NULL)   " & vbCrLf    '-- ��ҿ���
                SQL = SQL & "   AND (L.LABRESULT = ''  OR L.LABRESULT IS NULL)  " & vbCrLf
                SQL = SQL & "   AND L.LABENDDEP < '3'                           " & vbCrLf    '-- ó������ (2:����, 3:����Է�)
                    
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

'-- �˻��� ITEM ��������
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
    
    GetSampleITEM = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    strInOut = Trim(GetText(SPD, asRow, colINOUT))
    
    If strBarcode = "" Then
        Exit Function
    End If
        
    Select Case gEMR
        Case "HCMS"
            SQL = ""
            SQL = SQL & "SELECT IV_EXMN_CD              AS ITEM                 " & vbCrLf
            SQL = SQL & "  FROM LIME.VW_READ_LSMACIF                            " & vbCrLf
            SQL = SQL & " WHERE SPCM_RCPN_NO            = '" & strBarcode & "'  " & vbCrLf
            If strRegDate <> "" Then
                SQL = SQL & "   AND SUPT_DEPT_RCPN_DT   = '" & strRegDate & "'  " & vbCrLf
            End If
            SQL = SQL & "   AND LAB_MAC_CODE            = '" & gHOSP.PARTCD & "'" & vbCrLf
            SQL = SQL & "   AND IV_EXMN_CD              IN (" & gAllTestCd & ") " & vbCrLf
        
        Case "AMIS"
            SQL = ""
            SQL = SQL & "SELECT R.RESULTITEMCODE as ITEM                    " & vbCr
            SQL = SQL & "  FROM registinfos O, resultofnum R                " & vbCr
            SQL = SQL & " WHERE O.acptdate = R.acptdate                     " & vbCr
            SQL = SQL & "   AND R.SPCMNO = '" & strBarcode & "'             " & vbCr
            SQL = SQL & "   AND O.patid = R.patid                           " & vbCr
            SQL = SQL & "   AND O.acptseq = R.acptseq                       " & vbCr
            SQL = SQL & "   AND O.CLAS = 4                                  " & vbCr '�ӻ󺴸�
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
        
        Case "BIT"
            SQL = ""
            SQL = SQL & " SELECT DISTINCT R.ResLabCod AS ITEM                   " & vbCr
            SQL = SQL & "   FROM RESINF AS R                                    " & vbCr
            SQL = SQL & " WHERE LTRIM(RTRIM(R.RESOCMNUM)) = '" & strBarcode & "'" & vbCr
            SQL = SQL & "   AND R.RESLABCOD IN (" & gAllTestCd & ")             " & vbCr
            SQL = SQL & "   AND (R.RESREPTYP IS NULL OR R.RESREPTYP <> 'F')     " & vbCr         '--  'I':�߰� 'F' �Ϸ�"
            SQL = SQL & "   AND (R.RESRLTVAL = ''  OR R.RESRLTVAL IS NULL)      " & vbCr
            SQL = SQL & " Order By R.ResLabCod                                  " & vbCr
        
        Case "BIT70"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT L.LABODRCOD as ITEM                " & vbCrLf
            SQL = SQL & "  FROM ME_LABDAT L, ME_DAT D                       " & vbCrLf
            SQL = SQL & " WHERE L.LABODRDTE  = '" & strRegDate & "'         " & vbCrLf
            SQL = SQL & "   AND L.LABCHTNUM  = '" & strChartNo & "'         " & vbCrLf
            SQL = SQL & "   AND L.LABKEYNUM  = D.DATKEYNUM                  " & vbCrLf    '-- ���̺���Ű��
            SQL = SQL & "   AND L.LABATTEND  = D.DATATTEND                  " & vbCrLf    '-- ������ȣ
            SQL = SQL & "   AND L.LABATTEND  = M.MANATTEND                  " & vbCrLf    '-- ������ȣ ???
            SQL = SQL & "   AND L.LABCHTNUM  = D.DATCHTNUM                  " & vbCrLf    '-- íƮ��ȣ
            SQL = SQL & "   AND L.LABCHTNUM  = M.MANCHTNUM                  " & vbCrLf    '-- íƮ��ȣ
            SQL = SQL & "   AND L.LABODRDTE  = D.DATODRDTE                  " & vbCrLf    '-- ó������
            SQL = SQL & "   AND L.LABODRCOD IN (" & gAllTestCd & ")         " & vbCrLf
            SQL = SQL & "   AND (L.LABCANCEL = '' OR L.LABCANCEL IS NULL)   " & vbCrLf    '-- ��ҿ���
            SQL = SQL & "   AND (L.LABRESULT = ''  OR L.LABRESULT IS NULL)  " & vbCrLf
            SQL = SQL & "   AND L.LABENDDEP < '3'                           " & vbCrLf    '-- ó������ (2:����, 3:����Է�)
'            SQL = SQL & " Order By L.LABODRCOD                              " & vbCrLf
        
        Case "EONM"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT O.H141_SUGACD AS ITEM              " & vbCrLf
            SQL = SQL & "  FROM TB_H141_LISTAKEBODY O, TB_A110_PATINFO P    " & vbCrLf
            SQL = SQL & " Where P.A110_ChartNo = O.H141_CHARTNO             " & vbCrLf
            SQL = SQL & "   AND O.H141_TSAMPLENO  = '" & strBarcode & "'    " & vbCrLf
            'SQL = SQL & "   AND O.H141_NOTYYN = 'N'                         " & vbCrlf
            SQL = SQL & "   AND O.H141_NOTYYN       IN ('N','T')                 " & vbCrLf '������:T
            SQL = SQL & "   And O.H141_SUGACD in (" & gAllTestCd & ")       " & vbCrLf
            SQL = SQL & " ORDER BY O.H141_SUGACD                            " & vbCrLf
        
         Case "EASYS"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT ORD_CD AS ITEM                     " & vbCrLf
            SQL = SQL & "  FROM H3LAB_RESULT a, H1OPDIN b, HZ_MST_PTNT c    " & vbCrLf
            SQL = SQL & " WHERE a.ACC_YMD   = '" & strRegDate & "'          " & vbCrLf
            SQL = SQL & "   AND a.RECEPT_NO = '" & strBarcode & "'          " & vbCrLf
            SQL = SQL & "   AND a.ORD_CD IN (" & gAllTestCd & ")            " & vbCrLf
            SQL = SQL & "   AND a.STS_CD    = 'A'                           " & vbCrLf 'A:����, R:�������
            SQL = SQL & "   AND a.SUTAK_CD  = ''                            " & vbCrLf
            SQL = SQL & "   AND a.RECEPT_NO = b.RECEPT_NO                   " & vbCrLf
            SQL = SQL & " ORDER BY ORD_CD                                   " & vbCrLf
        
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
'            SQL = SQL & "    AND L.�˻����� = '" & gHOSP.LABCD & "'" & vbCrlf
            SQL = ""
            SQL = SQL & "SELECT DISTINCT (L.ó���ڵ� + L.�����ڵ�) AS ITEM                  " & vbCrLf
            SQL = SQL & "  FROM             TB_����˻� L                                   " & vbCrLf
            SQL = SQL & "       INNER JOIN  TB_�������� J ON  (L.��������ID = J.��������ID) " & vbCrLf
            SQL = SQL & "       INNER JOIN  TB_�����Ϲ� A ON  (J.��������   = A.��������    " & vbCrLf
            SQL = SQL & "                                AND   J.íƮ��ȣ   = A.íƮ��ȣ    " & vbCrLf
            SQL = SQL & "                                AND   J.�����ȣ   = A.�����ȣ)   " & vbCrLf
            SQL = SQL & " Where L.��ü��ȣ= '" & strBarcode & "'                            " & vbCrLf
            SQL = SQL & "   AND L.�˻���� < 5                                              " & vbCrLf
            SQL = SQL & "   AND L.ó���ڵ� + L.�����ڵ� IN (" & gAllTestCd & ")             " & vbCrLf
'            SQL = SQL & " ORDER BY L.ó���ڵ�, L.�����ڵ�                                   " & vbCrLf
            
        Case "KYU"
            SQL = ""
            
        
        Case "SANSOFT"
            SQL = ""
            SQL = SQL & "SELECT DISTINCT "
            SQL = SQL & "       E.�˻��ڵ�                              AS ITEM     " & vbCrLf
            SQL = SQL & "  FROM VW_�˻����� M, VW_�˻��� R, VW_�˻��ڵ� E         " & vbCrLf
            SQL = SQL & " WHERE M.�������� = CONVERT(DATE, '" & strRegDate & "')    " & vbCrLf
            SQL = SQL & "   AND M.�������� = R.��������                             " & vbCrLf
            SQL = SQL & "   AND M.������ȣ = R.������ȣ                             " & vbCrLf
            SQL = SQL & "   AND R.�˻��ڵ� = E.�˻��ڵ�                             " & vbCrLf
            SQL = SQL & "   AND m.������ȣ = '" & strPatID & "'                     " & vbCrLf
            SQL = SQL & "   AND E.�к��ڵ� = '" & gHOSP.PARTCD & "'                 " & vbCrLf
            SQL = SQL & "   AND E.�˻��ڵ� IN (" & gAllTestCd & ")                  " & vbCrLf
            SQL = SQL & "   AND IsNull(R.������, 'N') <> 'Y'                      " & vbCrLf
            SQL = SQL & "   AND (R.����� is null or R.����� = '')                 " & vbCrLf
        
        Case "LABSPEAR" 'PHILL
            SQL = ""
            SQL = SQL & "SELECT DISTINCT "
            SQL = SQL & "       E.�˻��ڵ�                              AS ITEM     " & vbCrLf
            SQL = SQL & "  FROM VW_�˻����� M, VW_�˻��� R, VW_�˻��ڵ� E         " & vbCrLf
            SQL = SQL & " WHERE M.�������� = CONVERT(DATE, '" & strRegDate & "')    " & vbCrLf
            SQL = SQL & "   AND M.�������� = R.��������                             " & vbCrLf
            SQL = SQL & "   AND M.������ȣ = R.������ȣ                             " & vbCrLf
            SQL = SQL & "   AND R.�˻��ڵ� = E.�˻��ڵ�                             " & vbCrLf
            SQL = SQL & "   AND m.������ȣ = '" & strPatID & "'                     " & vbCrLf
            SQL = SQL & "   AND E.�к��ڵ� = '" & gHOSP.PARTCD & "'                 " & vbCrLf
            SQL = SQL & "   AND E.�˻��ڵ� IN (" & gAllTestCd & ")                  " & vbCrLf
            SQL = SQL & "   AND IsNull(R.������, 'N') <> 'Y'                      " & vbCrLf
            SQL = SQL & "   AND (R.����� is null or R.����� = '')                 " & vbCrLf
            
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
            SQL = SQL & "Select DISTINCT (a.ó���ڵ� + a.�����ڵ�)      AS ITEM     " & vbCrLf
            SQL = SQL & "  From TB_�˻��׸� a, TB_����⺻ c                        " & vbCrLf
            SQL = SQL & " Where a.íƮ��ȣ = '" & strChartNo & "'                   " & vbCrLf
            SQL = SQL & "   And a.ó���ȣ > 0                                      " & vbCrLf
            SQL = SQL & "   And c.������� IN ('1','5','6','7','8','9')             " & vbCrLf
            SQL = SQL & "   And (a.ó���ڵ� + a.�����ڵ�) IN (" & gAllTestCd & ")   " & vbCrLf
            SQL = SQL & "   And (a.�˻��� IS NULL OR a.�˻��� = '')             " & vbCrLf
            SQL = SQL & "   And a.�����    = c.�����                              " & vbCrLf
            SQL = SQL & "   And a.�����    = c.�����                              " & vbCrLf
            SQL = SQL & "   And a.������    = c.������                              " & vbCrLf
            SQL = SQL & "   And a.íƮ��ȣ  = c.íƮ��ȣ                            " & vbCrLf
            SQL = SQL & "   And (a.�˻��� IS NULL OR a.�˻��� = '')             " & vbCrLf
            SQL = SQL & " Order By ITEM                                             " & vbCrLf

'            SQL = ""
'            SQL = SQL & "Select DISTINCT (a.ó���ڵ� + a.�����ڵ�)      AS ITEM     " & vbCrlf
'            SQL = SQL & "  from tb_�˻��׸� " & vbCrlf
'            SQL = SQL & " Where íƮ��ȣ = '" & argPID & "'" & vbCrlf
'            SQL = SQL & "   And �����   = '" & strYear & "'" & vbCrlf
'            SQL = SQL & "   And �����   = '" & strMonth & "'" & vbCrlf
'            SQL = SQL & "   And ������   = '" & strDay & "'" & vbCrlf
'            SQL = SQL & "   And ó���ȣ > 0 " & vbCrlf
'            SQL = SQL & "   And (�˻��� is null or �˻��� = '') " & vbCrlf
'            SQL = SQL & "   And ó���ڵ�+�����ڵ� in (" & gAllExam & ")"
        
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
            SQL = SQL & "   And OKFL <> 'Y'                 " & vbCrLf   '-- ���Ȯ������
            SQL = SQL & " Order By ORCD                     " & vbCrLf
        
        Case "NEOSOFT"
            If strInOut = "�Կ�" Then
                SQL = ""
                SQL = SQL & "SELECT DISTINCT a.CODE as ITEM                         " & vbCr
                SQL = SQL & "  From E_ORDER..ORDER_IN" & Format(Now, "yyyy") & " a  " & vbCr
                SQL = SQL & " Where a.CHAM_INDEX =  '" & strBarcode & "'            " & vbCr
                SQL = SQL & "   AND a.CODE IN (" & gAllTestCd & ")                  " & vbCr
                SQL = SQL & "   AND a.TRANS = '2'                                   " & vbCr
                SQL = SQL & " ORDER BY a.CODE                                       " & vbCr
            ElseIf strInOut = "�ܷ�" Then
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
            SQL = SQL & "   And B.EQUCODE1 = '" & gHOSP.MACHCD & "'     " & vbCr '����ڵ�
            SQL = SQL & "   AND A.MASTERCODE IN (" & gAllTestCd & ")    " & vbCr
            SQL = SQL & "   AND C.STATUS   <= '3'                       " & vbCr '�˻����
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
        
            Call SetSQLData("ITEM��ȸ", SQL)
    
            '-- Record Count ������
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
            
                
    Call SetSQLData("ITEM��ȸ", SQL)
    
    If SQL <> "" Then
        
        gPatOrdCd = ""
        
        '-- Record Count ������
        AdoCn.CursorLocation = adUseClient
        Set RS = AdoCn.Execute(SQL, , 1)
        If Not RS.EOF = True And Not RS.BOF = True Then
            Do Until RS.EOF
                If strItems = "" Then
                    strItems = GetTestNm(Trim(RS.Fields("ITEM")) & "", False)
                Else
                    strItems = strItems & "/" & GetTestNm(Trim(RS.Fields("ITEM")), False)
                End If
                
                'ó�� �˻������� �����´�.
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

'-- ����� ��ȸ
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
    SQL = SQL & "SELECT DISTINCT SAVESEQ,EXAMDATE,EXAMTIME,HOSPDATE,BARCODE,PNAME,SENDFLAG,SENDDATE " & vbCr
    '-- �˻���
    SQL = SQL & ",SEQNO,EXAMNAME,RESULT,PREVRESULT,REFJUDGE" & vbCr

    SQL = SQL & ",CHARTNO, PID,EQPSEQNO " & vbCr

    SQL = SQL & "  FROM PATRESULT " & vbCr
    '-- �˻�����
    SQL = SQL & " WHERE EXAMDATE Between '" & pFrom & "' AND '" & pTo & "'" & vbCr
'    SQL = SQL & "   AND EXAMCODE IN (" & gAllTestCd & ") " & vbCr
    SQL = SQL & " ORDER BY EXAMDATE,SAVESEQ,BARCODE,SEQNO"

    '-- Record Count ������
    AdoCn_Local.CursorLocation = adUseClient
    Set RS = AdoCn_Local.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        strItems = ""
        Do Until RS.EOF
            iCnt = iCnt + 1
            With SPD
                .ReDraw = False

                strSaveSeq = GetText(SPD, intRow, colSAVESEQ)
                strExamDate = GetText(SPD, intRow, colEXAMDATE)

                If strSaveSeq <> Trim(RS.Fields("SAVESEQ")) & "" Or strExamDate <> Trim(RS.Fields("EXAMDATE")) & "" Then
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows

                    SetText SPD, "1", intRow, colCHECKBOX
                    SetText SPD, Trim(RS.Fields("SAVESEQ")) & "", intRow, colSAVESEQ
                    SetText SPD, Trim(RS.Fields("EXAMDATE")) & "", intRow, colEXAMDATE
                    SetText SPD, Trim(RS.Fields("EXAMTIME")) & "", intRow, colEXAMTIME
                    SetText SPD, Trim(RS.Fields("HOSPDATE")) & "", intRow, colHOSPDATE
                    SetText SPD, Trim(RS.Fields("BARCODE")) & "", intRow, colBARCODE
                    SetText SPD, Trim(RS.Fields("CHARTNO")) & "", intRow, colCHARTNO
                    SetText SPD, Trim(RS.Fields("PID")) & "", intRow, colPID
                    SetText SPD, Trim(RS.Fields("PNAME")) & "", intRow, colPNAME
                    SetText SPD, Trim(RS.Fields("EQPSEQNO")) & "", intRow, colSEQNO


                    Select Case Trim(RS.Fields("SENDFLAG")) & ""
                    Case "0"
                            SetText SPD, "�����", intRow, colSTATE
                    Case "1"
                            SetText SPD, "���忡��", intRow, colSTATE
                    Case "2"
                            SetText SPD, "���ۿϷ�", intRow, colSTATE
                    End Select

                    'If gEMR <> "KOMAIN" Then
                    '    SetText SPD, GetSampleITEM(intRow, SPD), intRow, colITEMS
                    'End If
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
        'frmMain.chkRAll.Value = "1"
    Else
        'frmMain.lblStatus.Caption = ">> ��ȸ ����ڰ� �����ϴ�."
        'frmMain.chkRAll.Value = "0"
    End If

    RS.Close

    SPD.RowHeight(-1) = 15
    SPD.ReDraw = True

'    Call frmMain.GetPatTRestResult_Search(1)

    Screen.MousePointer = 0

End Sub

'-- �˻��� ��������
Function SaveTransData(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    
    SaveTransData = -1
    
    Select Case gEMR
        Case "KHNMC"
            SaveTransData = SaveTransData_KHNMC(argSpcRow, SPD)
        
'        Case "SUNGMO"
'            SaveTransData = SaveTransData_SUNGMO(argSpcRow, SPD)
        
        Case "NU"
            SaveTransData = SaveTransData_NU(argSpcRow, SPD)
        
        Case "HCMS"
            SaveTransData = SaveTransData_HCMS(argSpcRow, SPD)
        
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
        
        Case "MCC"
            SaveTransData = SaveTransData_MCC(argSpcRow, SPD)

        Case "TWIN"
            SaveTransData = SaveTransData_TWIN(argSpcRow, SPD)
    End Select


End Function

'-- �˻��� ��������
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
'   ��� : �ش� ���ڵ��ȣ�� ���� 1. �������� ��ȸ,
'                                 2. ���������� ȭ��ǥ��,
'                                 3. ó���ڵ� ��������
'   �μ� :
'       - pBarNo : ���ڵ��ȣ
'       - pType  : ���ڵ� �̻��� ���ϴ� ���
'                   1 : Seq
'                   2 : Rack/Pos
'                   3 : üũ�Ȱ��� ���� ���� ��
'-----------------------------------------------------------------------------'
Public Sub SetPatInfo(ByVal pBarNo As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    

    intRow = -1
    With frmMain
        Select Case pType
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmMain.spdOrder, i, colBARCODE) = pBarNo Then
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmMain.spdOrder, i, colSEQNO) = mResult.Seq Then
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                        intRow = i
                        Exit For
                    End If
                Next i
        End Select
        
        '-- �������忡�� ��ã����..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If
    
        
        '-- ������ε��� ȭ��ǥ��
        Call SetText(.spdOrder, "1", intRow, colCHECKBOX)
        Call SetText(.spdOrder, mResult.RsltDate, intRow, colEXAMDATE)
        Call SetText(.spdOrder, mResult.RsltTime, intRow, colEXAMTIME)
        Call SetText(.spdOrder, mResult.RsltSeq, intRow, colSAVESEQ)
        'If gHOSP.BARUSE = "Y" Then
        '    Call SetText(.spdOrder, mResult.BarNo, intRow, colBARCODE)
        'End If
        'Call SetText(.spdOrder, mResult.RackNo, intRow, colRACKNO)
        'Call SetText(.spdOrder, mResult.TubePos, intRow, colPOSNO)
        Call SetText(.spdOrder, mResult.Seq, intRow, colSEQNO)
    
        '-- ����������� �����
        .spdResult.MaxRows = 0
    
        'If gHOSP.BARUSE = "Y" Then
        
            'Call XmlSelect_Free
        
            '-- �˻��� ���� ��������
            Call GetSampleInfo(intRow, .spdOrder)
        
        'End If
        
        .spdOrder.RowHeight(-1) = 15
        
    End With
    
    '-- ���� Row
    gRow = intRow
    
End Sub

Public Sub XmlSelect_Free()
    
    With XmlSelect
        .AGE = ""
        .BCNO = ""
        .EXECprcpuniqno = ""
        .ORDDEPTCD = ""
        .PATNM = ""
        .PID = ""
        .PRCPDD = ""
        .PRGSTNO = ""
        .RETESTYN = ""
        .RSLTSTAT = ""
        .SEX = ""
        .SPCACPTDT = ""
        .SPCCD = ""
        .SPCNM = ""
        .SPCSTAT = ""
        .TCLSCD = ""
        .TestCd = ""
        .TESTLRGCD = ""
        .WORKNO = ""
    End With
    
End Sub


'-- �˻��� ���� ��������
Function GetSampleInfo(ByVal asRow As Long, ByVal SPD As Object) As Integer

    Screen.MousePointer = 11

    GetSampleInfo = -1

    'If cn_Server_Flag = True Then
        Select Case gEMR
            Case "KHNMC"
                    Call GetSampleInfo_KHNMC(asRow, SPD)

            Case "NU"
                    Call GetSampleInfo_NU(asRow, SPD)
            
            Case "HCMS"
                    Call GetSampleInfo_HCMS(asRow, SPD)
                    
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
            
            Case "MCC"
                    Call GetSampleInfo_MCC(asRow, SPD)
    
            Case "TWIN"
                    Call GetSampleInfo_TWIN(asRow, SPD)
    
    
        End Select
    
        GetSampleInfo = 1
    
    'End If
    
    Screen.MousePointer = 0


End Function

'-- �˻��� ���� ��������
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
        
    Call SetSQLData("���ڵ���ȸ", SQL)
    
    '-- Record Count ������
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
                
                '��������
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '���������� ����
                With mOrder
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                'ȯ�� ����/����
                With mPatient
                    .AGE = Trim(RS.Fields("AGE")) & ""
                    .SEX = Trim(RS.Fields("SEX")) & ""
                End With
                
                '-- ȭ�鿡 ǥ��
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "��", asRow, intCol)
                        '-- ó���ڵ�
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "GetSampleInfo_PHILL" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function


'-- �˻��� ���� ��������
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
    SQL = SQL & "   AND O.H141_NOTYYN       IN ('N','T')                 " & vbCr '������:T
    SQL = SQL & "   And O.H141_SUGACD in (" & gAllTestCd & ")     " & vbCrLf
    SQL = SQL & " Order By O.H141_SUGACD                          " & vbCrLf
        
    Call SetSQLData("���ڵ���ȸ", SQL)
    
    '-- Record Count ������
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
                
                '��������
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '���������� ����
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
'                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                'ȯ�� ����/����
                With mPatient
                    .AGE = Trim(RS.Fields("AGE")) & ""
                    .SEX = Trim(RS.Fields("SEX")) & ""
                End With
                
                '-- ȭ�鿡 ǥ��
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "��", asRow, intCol)
                        
                        '-- �������� SEQ
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "GetSampleInfo_EONM" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- �˻��� ���� ��������
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
    SQL = SQL & "   AND O.CLAS          = 4                             " & vbCrLf '�ӻ󺴸�
    SQL = SQL & "   AND R.SPCMNO        = '" & strBarcode & "'          " & vbCrLf
    SQL = SQL & "   AND (R.NUMRESULTVAL = '' OR R.NUMRESULTVAL IS NULL) " & vbCrLf
    SQL = SQL & "   AND R.RESULTITEMCODE IN (" & gAllTestCd & ")        " & vbCrLf
        
    Call SetSQLData("���ڵ���ȸ", SQL)
    
    '-- Record Count ������
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
                
                '��������
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '���������� ����
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                'ȯ�� ����/����
                With mPatient
                    .SEX = Trim(RS.Fields("SEX")) & ""
                    '.AGE = Trim(RS.Fields("AGE")) & ""
                End With
                
                '-- ȭ�鿡 ǥ��
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "��", asRow, intCol)
                        
                        '-- ó���ڵ�
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_GetSampleInfo_AMIS" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- �˻��� ���� ��������
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
    SQL = SQL & "       J.��������                  AS HOSPDATE                 " & vbCrLf
    SQL = SQL & "     , L.��ü��ȣ                  AS BARCODE                  " & vbCrLf
    SQL = SQL & "     , A.íƮ��ȣ                  AS CHARTNO                  " & vbCrLf
    SQL = SQL & "     , J.������ȣ                  AS PID                      " & vbCrLf
    SQL = SQL & "     , A.ȯ���̸�                  AS PNAME                    " & vbCrLf
    SQL = SQL & "     , A.ȯ�ڼ���                  AS SEX                      " & vbCrLf
    SQL = SQL & "     , A.ȯ�ڳ���                  AS AGE                      " & vbCrLf
    SQL = SQL & "     , L.����˻�ID                AS TESTID                   " & vbCrLf
    SQL = SQL & "     , L.��������ID                AS SPRTID                   " & vbCrLf
    SQL = SQL & "     , (L.ó���ڵ�+ L.�����ڵ�)    AS ITEM                     " & vbCrLf
    SQL = SQL & "  FROM         TB_����˻� L                                   " & vbCrLf
    SQL = SQL & "   INNER JOIN  TB_�������� J ON (L.��������ID = J.��������ID)  " & vbCrLf
    SQL = SQL & "   INNER JOIN  TB_�����Ϲ� A ON (J.��������   = A.��������     " & vbCrLf
    SQL = SQL & "                            AND  J.íƮ��ȣ   = A.íƮ��ȣ     " & vbCrLf
    SQL = SQL & "                            AND  J.�����ȣ   = A.�����ȣ)    " & vbCrLf
    SQL = SQL & " Where L.��ü��ȣ = '" & strBarcode & "'                       " & vbCrLf
    SQL = SQL & "   AND L.�˻���� < 5                                          " & vbCrLf
    SQL = SQL & "   AND L.ó���ڵ� + L.�����ڵ� IN (" & gAllTestCd & ")         " & vbCrLf
    SQL = SQL & " ORDER BY J.��������, L.��ü��ȣ                               " & vbCrLf
    Call SetSQLData("���ڵ���ȸ", SQL)
    
    '-- Record Count ������
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
                
                '��������
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '���������� ����
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                'ȯ�� ����/����
                With mPatient
                    .SEX = Trim(RS.Fields("SEX")) & ""
                    .AGE = Trim(RS.Fields("AGE")) & ""
                End With
                
                '-- ȭ�鿡 ǥ��
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "��", asRow, intCol)
                        
                        '-- ����˻�ID
                        gArrEQP(intCol - colSTATE, 16) = Trim(RS.Fields("TESTID")) & ""
                        
                        '-- ��������ID
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_GetSampleInfo_KCHART" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- �˻��� ���� ��������
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
    
    If Not IsNumeric(strBarcode) Then
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
    
    If frmMain.chkSave.Value = "0" Then
        SQL = SQL & "   AND a.JSTATUS < '3'                     " & vbCrLf
    End If
    
    SQL = SQL & " ORDER BY b.LABCODE                        " & vbCrLf
    
    Call SetSQLData("���ڵ���ȸ", SQL)
    
    
    '-- Record Count ������
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
                
                '��������
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '���������� ����
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                
                '-- ȭ�鿡 ǥ��
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "��", asRow, intCol)
                        
                        '-- ó���ڵ�
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_GetSampleInfo_JWINFO" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- �˻��� ���� ��������
Function GetSampleInfo_NU(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    
    Dim i           As Integer
    Dim J           As Integer
    Dim sParam      As String
    Dim sRcvData    As String
    Dim varRcvData  As Variant
    Dim varTstCode  As Variant
    Dim strXmlName  As String
    Dim strNames    As String
    
    
On Error GoTo DBErr
    
    GetSampleInfo_NU = -1
    
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
    


    '���ٽ�
    'string sendString = SERVERIP + "/himed2/.live?submit_id=TRLII00123&business_id=lis&instcd=053&bcno=" + barcode;
'    sParam = ""
'    sParam = sParam & "submit_id=TRLII00123&"                                   'submit ID
'    sParam = sParam & "business_id=lis&"                                        'business_id
'    sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"                            '����ڵ�
'    sParam = sParam & "bcno=" & strBarcode                                      '�˻��ڵ�
    
    
    '�˻��ڵ� ��ȸ
    'http://his202.sejongh.co.kr/himed/webapps/com/commonweb/xrw/.live?submit_id=TRLII00103&business_id=lis&ex_interface=19920017|202&instcd=202&eqmtcd=O02&refgbn=2

'    sParam = ""
'    sParam = sParam & "submit_id=TRLII00101&"                                        'submit ID
'    sParam = sParam & "business_id=lis&"                                    'business_id
'    sParam = sParam & "bcno=" & strBarcode & "&"                                                        '���ڵ�
'    sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"                           '����ڵ�
'    sParam = sParam & "eqmtcd=" & gHOSP.MACHCD & "&"                           '����ڵ�
    
    sParam = ""
    sParam = sParam & "submit_id=TRLII00101&"                                        'submit ID
    sParam = sParam & "business_id=li&"                                    'business_id
    sParam = sParam & "ex_interface=" & gHOSP.USERID & "|" & gHOSP.HOSPCD & "&"      '�����ID|����ڵ�
    sParam = sParam & "instcd=" & gHOSP.HOSPCD & "&"                           '����ڵ�
    sParam = sParam & "eqmtcd=" & gHOSP.MACHCD & "&"                           '����ڵ�
    sParam = sParam & "bcno=" & strBarcode                                                        '���ڵ�
    
    
    sRcvData = OpenURLWithIE2(gHOSP.APIURL & sParam, frmMain.Inet1)

    Call SetSQLData("���ڵ���ȸ", "Param:" & gHOSP.APIURL & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine, "A")
        
    If InStr(1, sRcvData, "<?xml version") > 0 Then
        varRcvData = Split(sRcvData, "<worklist>")
    End If
    
    strXmlName = gHOSP.MACHNM & "_" & Format(CDate(Now), "yyyymmdd") & "_" & strBarcode & ".xml"
    
    Call SetXMLData(strXmlName, sRcvData)
    
    Call DisplayNode_InfoS(App.PATH & "\Xml\" & strXmlName, UBound(varRcvData))
    'Call DisplayNode_InfoS("D:\������Ʈ\VB\__NU\��õ���𺴿�_NU\����\patlist2.xml", 1)

    'Kill App.PATH & "\Xml\" & strXmlName


'    If UBound(varRcvData) >= 0 Then
    If UBound(varRcvData) >= 0 Then
        For i = 0 To UBound(varRcvData) - 1
        'For i = 0 To 0
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1

                'SetText SPD, "1", asRow, colCHECKBOX
                'SetText SPD, XmlSelectS.SPCACPTDT(i), asRow, colHOSPDATE
                'SetText SPD, XmlSelectS.BCNO(i), asRow, colBARCODE
                'SetText SPD, XmlSelectS.PID(i), asRow, colPID
                'SetText SPD, XmlSelectS.PATNM(i), asRow, colPNAME
                'SetText SPD, XmlSelectS.SEX(i), asRow, colPSEX
                'SetText SPD, XmlSelectS.AGE(i), asRow, colPAGE
                'SetText SPD, XmlSelectS.SPCNM(i), asRow, colSPECIMEN
                                                                 
                '���������� ����
'                With mOrder
'                    .BarNo = XmlSelectS.BCNO(i)
'                    .PID = XmlSelectS.PID(i)
'                    .PNAME = XmlSelectS.PATNM(i)
'                    .Count = CStr(intTestCnt)
'                    .NoOrder = False
'                End With
    
                varTstCode = Split(XmlSelectS.TCLSCDLIST(i), "��")
                For J = 0 To UBound(varTstCode) - 1
                    
                    '-- ȭ�鿡 ǥ��
                    For intCol = colSTATE + 1 To .MaxCols
                        If GetTestNm(Trim(varTstCode(J))) = gArrEQPNm(intCol - colSTATE, 6) Then
                            .Row = asRow
                            .Col = intCol
                            .BackColor = vbYellow
                            Call SetText(SPD, "��", asRow, intCol)
                            Exit For
                        End If
                    Next
                    
                    gPatOrdCd = gPatOrdCd & "'" & Trim(varTstCode(J)) & "',"
                Next
            End With
        Next
    End If
    
    If gPatOrdCd <> "" Then
        gPatOrdCd = gPatOrdCd & "'XXXXX',"
        gPatOrdCd = gPatOrdCd & "'YYYYY',"
        gPatOrdCd = gPatOrdCd & "'ZZZZZ',"
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    
    GetSampleInfo_NU = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_NU = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
End Function

'-- �˻��� ���� ��������
Function GetSampleInfo_KHNMC(ByVal asRow As Long, ByVal SPD As vaSpread) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    
    Dim i           As Integer
    Dim J           As Integer
    Dim sParam      As String
    Dim sRcvData    As String
    Dim varRcvData  As Variant
    Dim varRcvDtl   As Variant
    Dim varTstCode  As Variant
    Dim strXmlName  As String
    Dim strNames    As String
    Dim strTstCode  As String
    
On Error GoTo DBErr
    
    GetSampleInfo_KHNMC = -1
    
    intTestCnt = 0
    
    gPatOrdCd = ""
    strTstCode = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
    
    'http://��������IP/lis/jindangeomsaweb/GyeolGwaIF.live?Mode=reqGetMWLBCDJoHoi&Data1=���ڵ�

    sParam = ""
    sParam = sParam & "Data1=" & strBarcode     '���ڵ�
    
    sRcvData = OpenURLWithIE2("http://" & gHOSP.APIURL & "/lis/jindangeomsaweb/GyeolGwaIF.live?Mode=reqGetMWLBCDJoHoi&" & sParam, frmMain.Inet1)

    Call SetSQLData("���ڵ���ȸ", "http://" & gHOSP.APIURL & "/lis/jindangeomsaweb/GyeolGwaIF.live?Mode=reqGetMWLBCDJoHoi&" & sParam & vbNewLine & "Return:" & sRcvData & vbNewLine, "A")
        
'    MsgBox sRcvData
    
    If InStr(1, sRcvData, "<?xml version") > 0 Then
        varRcvData = Split(sRcvData, "<IFordcd>")
    End If
    
    'strXmlName = gHOSP.MACHNM & "_" & Format(CDate(Now), "yyyymmdd") & "_" & strBarcode & ".xml"
    'Call SetXMLData(strXmlName, sRcvData)
    'Call DisplayNode_InfoS(App.PATH & "\Xml\" & strXmlName, UBound(varRcvData))
    'Call DisplayNode_InfoS("D:\������Ʈ\VB\__NU\��õ���𺴿�_NU\����\patlist2.xml", 1)
    'Call DisplayNode_InfoS_Worklist_KHNMC("D:\������Ʈ\VB\��������뺴��\IF\XML\worklist.xml", 10)
    

    'Kill App.PATH & "\Xml\" & strXmlName
    
'    MsgBox UBound(varRcvData) ==> 1
    
    If UBound(varRcvData) >= 1 Then
        For i = 1 To UBound(varRcvData)

'MsgBox varRcvData(i)
            
            varRcvData(i) = Replace(varRcvData(i), vbLf, "")
            varRcvDtl = Split(varRcvData(i), vbCr)
            With SPD
                .ReDraw = False
                intTestCnt = intTestCnt + 1

                'varTstCode = Split(mGetP(varRcvDtl(i), 9, "��"))
                'For J = 0 To UBound(varRcvDtl) - 1
                For J = 1 To 1
                    
'MsgBox varRcvDtl(0)
'MsgBox varRcvDtl(1)
'MsgBox varRcvDtl(2)
'MsgBox varRcvDtl(3)
                    
                    strTstCode = mGetP(varRcvDtl(J), 13, "��")
                    '-- ȭ�鿡 ǥ��
                    For intCol = colSTATE + 1 To .MaxCols
                        If GetTestNm(strTstCode) = gArrEQPNm(intCol - colSTATE, 6) Then
                            .Row = asRow
                            .Col = intCol
                            .BackColor = vbYellow
                            Call SetText(SPD, "��", asRow, intCol)
                            Exit For
                        End If
                    Next
                    
                    'gPatOrdCd = gPatOrdCd & "'" & Trim(Split(mGetP(varRcvDtl(i), 9, "��"))) & "',"
                    gPatOrdCd = gPatOrdCd & "'" & strTstCode & "',"
                Next
            End With
        Next
    End If
    
    If gPatOrdCd <> "" Then
        gPatOrdCd = gPatOrdCd & "'XXXXX',"
        gPatOrdCd = gPatOrdCd & "'YYYYY',"
        gPatOrdCd = gPatOrdCd & "'ZZZZZ',"
        gPatOrdCd = Mid(gPatOrdCd, 1, Len(gPatOrdCd) - 1)
    End If
    
    
    GetSampleInfo_KHNMC = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_KHNMC = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
End Function

'ȯ������ ��ȸ
Public Sub DisplayNode_InfoS(asPath As String, asCnt As Integer)

    Dim xmlDoc          As New MSXML2.DOMDocument30
    Dim nodeBook        As IXMLDOMElement
    Dim nodeId          As IXMLDOMAttribute
    Dim xNode           As MSXML2.IXMLDOMNode
    Dim namedNodeMap    As IXMLDOMNamedNodeMap
    Dim Child_Node      As MSXML2.IXMLDOMNodeList
    Dim i               As Integer
    Dim J               As Integer
    Dim k               As Integer
    Dim intNodeLen      As Integer
    
On Error GoTo ErrXML:
    
    Set xmlDoc = New MSXML2.DOMDocument30
    
    xmlDoc.async = False
    xmlDoc.Load asPath
    
    If (xmlDoc.parseError.errorCode <> 0) Then
        Dim myErr
        Set myErr = xmlDoc.parseError
        MsgBox ("You have error " & myErr.reason)
    Else
        ReDim Preserve XmlSelectS.AGE(asCnt)
        ReDim Preserve XmlSelectS.BCNO(asCnt)
        ReDim Preserve XmlSelectS.EXECprcpuniqno(asCnt)
        ReDim Preserve XmlSelectS.ORDDEPTCD(asCnt)
        ReDim Preserve XmlSelectS.PATNM(asCnt)
        ReDim Preserve XmlSelectS.PID(asCnt)
        ReDim Preserve XmlSelectS.PRCPDD(asCnt)
        ReDim Preserve XmlSelectS.PRGSTNO(asCnt)
        ReDim Preserve XmlSelectS.RETESTYN(asCnt)
        ReDim Preserve XmlSelectS.RSLTSTAT(asCnt)
        ReDim Preserve XmlSelectS.SEX(asCnt)
        ReDim Preserve XmlSelectS.SPCACPTDT(asCnt)
        ReDim Preserve XmlSelectS.SPCCD(asCnt)
        ReDim Preserve XmlSelectS.SPCNM(asCnt)
        ReDim Preserve XmlSelectS.SPCSTAT(asCnt)
        ReDim Preserve XmlSelectS.TCLSCD(asCnt)
        ReDim Preserve XmlSelectS.TCLSCDLIST(asCnt)
        ReDim Preserve XmlSelectS.TestCd(asCnt)
        ReDim Preserve XmlSelectS.TESTLRGCD(asCnt)
        ReDim Preserve XmlSelectS.WORKNO(asCnt)
        ReDim Preserve XmlSelectS.IFREQCDLIST(asCnt)
        'SPCACPTDT
            
        '<bcno><![CDATA[3010700030]]></bcno>
        '<patnm><![CDATA[�ڼ���]]></patnm>
        '<prgstno><![CDATA[400321-1******]]></prgstno>
        '<pid><![CDATA[000132623]]></pid>
        '<sex><![CDATA[M]]></sex>
        '<age><![CDATA[78]]></age>
        '<spcnm><![CDATA[Throat swab]]></spcnm>
        '<spccd><![CDATA[023]]></spccd>
        '<tclscd><![CDATA[VB6012A]]></tclscd>
        '<spcstat><![CDATA[4]]></spcstat>
        '<rsltstat><![CDATA[-]]></rsltstat>
        '<workno><![CDATA[20181217I20002]]></workno>
        '<testcd><![CDATA[VB6012A]]></testcd>
        '<execprcpuniqno><![CDATA[2002638354]]></execprcpuniqno>
        '<spcacptdt><![CDATA[20181217094414]]></spcacptdt>
        '<prcpdd><![CDATA[20181217]]></prcpdd>
        '<retestyn><![CDATA[N]]></retestyn>
        '<testlrgcd><![CDATA[I]]></testlrgcd>
        '<orddeptcd><![CDATA[NU]]></orddeptcd>
        
        
        Set Child_Node = xmlDoc.childNodes
        For Each xNode In Child_Node
            If xNode.nodeType = NODE_ELEMENT Then
                For intNodeLen = 0 To xNode.childNodes.Length - 1
                    For i = 0 To xNode.childNodes.Item(intNodeLen).childNodes.Length - 1
                        For k = 0 To xNode.childNodes.Item(intNodeLen).childNodes.Item(i).childNodes.Length - 1
                            'Debug.Print xNode.childNodes.Item(intNodeLen).childNodes.Item(i).baseName & ":" & xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue
                            Debug.Print UCase(xNode.childNodes.Item(intNodeLen).childNodes.Item(i).childNodes.Item(k).baseName)
                            Select Case UCase(xNode.childNodes.Item(intNodeLen).childNodes.Item(i).childNodes.Item(k).baseName)
                                Case "AGE":             XmlSelectS.AGE(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).childNodes.Item(k).nodeTypedValue                  '����           [78]
                                Case "BCNO":            XmlSelectS.BCNO(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).childNodes.Item(k).nodeTypedValue                 '��ü��ȣ       [3010700030]
                                Case "EXECprcpuniqno":  XmlSelectS.EXECprcpuniqno(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).childNodes.Item(k).nodeTypedValue       'íƮ��ȣ?      [2002638354]
                                Case "ORDDEPTCD":       XmlSelectS.ORDDEPTCD(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).childNodes.Item(k).nodeTypedValue            'ó��μ��ڵ�?  [NU]
                                
                                Case "PATNM":           XmlSelectS.PATNM(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).childNodes.Item(k).nodeTypedValue                'ȯ�ڸ�         [�ڼ���]
                                Case "PID":             XmlSelectS.PID(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).childNodes.Item(k).nodeTypedValue                  'ȯ�ڹ�ȣ       [000132623]
                                Case "PRCPDD":          XmlSelectS.PRCPDD(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).childNodes.Item(k).nodeTypedValue               'ó����?        [20181217]
                                Case "PRGSTNO":         XmlSelectS.PRGSTNO(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).childNodes.Item(k).nodeTypedValue              '�ֹι�ȣ       [400321-1******]
                                
                                Case "RETESTYN":        XmlSelectS.RETESTYN(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).childNodes.Item(k).nodeTypedValue             '��˿���       [N]
                                Case "RSLTSTAT":        XmlSelectS.RSLTSTAT(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).childNodes.Item(k).nodeTypedValue             '�������       [-]
                                
                                Case "SEX":             XmlSelectS.SEX(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).childNodes.Item(k).nodeTypedValue                  '����           [M]
                                
                                Case "SPCACPTDT":       XmlSelectS.SPCACPTDT(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).childNodes.Item(k).nodeTypedValue           '��üä��ð�?  [20181217094414]
                                Case "SPCCD":           XmlSelectS.SPCCD(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).childNodes.Item(k).nodeTypedValue                '��ü�ڵ�       [023]
                                Case "SPCNM":           XmlSelectS.SPCNM(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).childNodes.Item(k).nodeTypedValue                '��ü��         [Throat swab]
                                Case "SPCSTAT":         XmlSelectS.SPCSTAT(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).childNodes.Item(k).nodeTypedValue              '��ü����       [4]
                                
                                Case "TCLSCD":          XmlSelectS.TCLSCD(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).childNodes.Item(k).nodeTypedValue               'ó���ڵ�       [VB6012A]
                                Case "TCLSCDLIST":      XmlSelectS.TCLSCDLIST(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).childNodes.Item(k).nodeTypedValue           '�˻��ڵ�       [CDATA[LPD329��]
                                Case "IFREQCDLIST":     XmlSelectS.IFREQCDLIST(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).childNodes.Item(k).nodeTypedValue                                                                                                                                                               ' [CDATA[LPD31103��LPD31112��]
                                Case "TESTCD":          XmlSelectS.TestCd(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).childNodes.Item(k).nodeTypedValue               '�˻��ڵ�       [VB6012A]
                                
                                Case "TESTLRGCD":       XmlSelectS.TESTLRGCD(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).childNodes.Item(k).nodeTypedValue            '����׷��ڵ�?  [I]
                                Case "WORKNO":          XmlSelectS.WORKNO(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).childNodes.Item(k).nodeTypedValue               '��ũ��ȣ       [20181217I20002]
                                
                                '1 : SPCACPTDT
                                '2 : ACPTFLAG
                                '3 : BCNO

                            End Select
                        Next
                        J = J + 1
                    Next
                    'J = J + 1
                Next
            End If
        Next
       
        Set Child_Node = Nothing
        
    End If

    Exit Sub
    
ErrXML:
    Exit Sub
    
End Sub

'��ũ����Ʈ ��ȸ
Public Sub DisplayNode_InfoS_Worklist(asPath As String, asCnt As Integer)

    Dim xmlDoc          As New MSXML2.DOMDocument30
    Dim nodeBook        As IXMLDOMElement
    Dim nodeId          As IXMLDOMAttribute
    Dim xNode           As MSXML2.IXMLDOMNode
    Dim namedNodeMap    As IXMLDOMNamedNodeMap
    Dim Child_Node      As MSXML2.IXMLDOMNodeList
    Dim i, J            As Integer
    Dim intNodeLen      As Integer
    
On Error GoTo ErrXML:
    
    Set xmlDoc = New MSXML2.DOMDocument30
    
    xmlDoc.async = False
    xmlDoc.Load asPath
    
    If (xmlDoc.parseError.errorCode <> 0) Then
        Dim myErr
        Set myErr = xmlDoc.parseError
        MsgBox ("You have error " & myErr.reason)
    Else
        ReDim Preserve XmlSelectS.AGE(asCnt)
        ReDim Preserve XmlSelectS.BCNO(asCnt)
        ReDim Preserve XmlSelectS.EXECprcpuniqno(asCnt)
        ReDim Preserve XmlSelectS.ORDDEPTCD(asCnt)
        ReDim Preserve XmlSelectS.PATNM(asCnt)
        ReDim Preserve XmlSelectS.PID(asCnt)
        ReDim Preserve XmlSelectS.PRCPDD(asCnt)
        ReDim Preserve XmlSelectS.PRGSTNO(asCnt)
        ReDim Preserve XmlSelectS.RETESTYN(asCnt)
        ReDim Preserve XmlSelectS.RSLTSTAT(asCnt)
        ReDim Preserve XmlSelectS.SEX(asCnt)
        ReDim Preserve XmlSelectS.SPCACPTDT(asCnt)
        ReDim Preserve XmlSelectS.SPCCD(asCnt)
        ReDim Preserve XmlSelectS.SPCNM(asCnt)
        ReDim Preserve XmlSelectS.SPCSTAT(asCnt)
        ReDim Preserve XmlSelectS.TCLSCD(asCnt)
        ReDim Preserve XmlSelectS.TestCd(asCnt)
        ReDim Preserve XmlSelectS.TESTLRGCD(asCnt)
        ReDim Preserve XmlSelectS.WORKNO(asCnt)
            
        '<bcno><![CDATA[3010700030]]></bcno>
        '<patnm><![CDATA[�ڼ���]]></patnm>
        '<prgstno><![CDATA[400321-1******]]></prgstno>
        '<pid><![CDATA[000132623]]></pid>
        '<sex><![CDATA[M]]></sex>
        '<age><![CDATA[78]]></age>
        '<spcnm><![CDATA[Throat swab]]></spcnm>
        '<spccd><![CDATA[023]]></spccd>
        '<tclscd><![CDATA[VB6012A]]></tclscd>
        '<spcstat><![CDATA[4]]></spcstat>
        '<rsltstat><![CDATA[-]]></rsltstat>
        '<workno><![CDATA[20181217I20002]]></workno>
        '<testcd><![CDATA[VB6012A]]></testcd>
        '<execprcpuniqno><![CDATA[2002638354]]></execprcpuniqno>
        '<spcacptdt><![CDATA[20181217094414]]></spcacptdt>
        '<prcpdd><![CDATA[20181217]]></prcpdd>
        '<retestyn><![CDATA[N]]></retestyn>
        '<testlrgcd><![CDATA[I]]></testlrgcd>
        '<orddeptcd><![CDATA[NU]]></orddeptcd>
        
        
        Set Child_Node = xmlDoc.childNodes
        For Each xNode In Child_Node
            If xNode.nodeType = NODE_ELEMENT Then
                For intNodeLen = 0 To xNode.childNodes.Length - 1
                    For i = 0 To xNode.childNodes.Item(intNodeLen).childNodes.Length - 1
                        'Debug.Print xNode.childNodes.Item(intNodeLen).childNodes.Item(i).baseName & ":" & xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue
                        Select Case UCase(xNode.childNodes.Item(intNodeLen).childNodes.Item(i).baseName)
                            Case "AGE":             XmlSelectS.AGE(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue                 '����           [78]
                            Case "BCNO":            XmlSelectS.BCNO(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue                '��ü��ȣ       [3010700030]
                            Case "EXECprcpuniqno":  XmlSelectS.EXECprcpuniqno(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue      'íƮ��ȣ?      [2002638354]
                            Case "ORDDEPTCD":       XmlSelectS.ORDDEPTCD(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue           'ó��μ��ڵ�?  [NU]
                            
                            Case "PATNM":           XmlSelectS.PATNM(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue               'ȯ�ڸ�         [�ڼ���]
                            Case "PID":             XmlSelectS.PID(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue                 'ȯ�ڹ�ȣ       [000132623]
                            Case "PRCPDD":          XmlSelectS.PRCPDD(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue              'ó����?        [20181217]
                            Case "PRGSTNO":         XmlSelectS.PRGSTNO(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue             '�ֹι�ȣ       [400321-1******]
                            
                            Case "RETESTYN":        XmlSelectS.RETESTYN(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue            '��˿���       [N]
                            Case "RSLTSTAT":        XmlSelectS.RSLTSTAT(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue            '�������       [-]
                            
                            Case "SEX":             XmlSelectS.SEX(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue                 '����           [M]
                            
                            Case "SPCACPTDT":       XmlSelectS.SPCACPTDT(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue           '��üä��ð�?  [20181217094414]
                            Case "SPCCD":           XmlSelectS.SPCCD(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue               '��ü�ڵ�       [023]
                            Case "SPCNM":           XmlSelectS.SPCNM(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue               '��ü��         [Throat swab]
                            Case "SPCSTAT":         XmlSelectS.SPCSTAT(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue             '��ü����       [4]
                            
                            Case "TCLSCD":          XmlSelectS.TCLSCD(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue              'ó���ڵ�       [VB6012A]
                            Case "TESTCD":          XmlSelectS.TestCd(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue              '�˻��ڵ�       [VB6012A]
                            Case "TESTLRGCD":       XmlSelectS.TESTLRGCD(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue           '����׷��ڵ�?  [I]
                            Case "WORKNO":          XmlSelectS.WORKNO(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue              '��ũ��ȣ       [20181217I20002]
                        End Select
                    Next
                    J = J + 1
                Next
            End If
        Next
       
        Set Child_Node = Nothing
        
    End If

    Exit Sub
    
ErrXML:
    Exit Sub
    
End Sub

'��ũ����Ʈ ��ȸ
Public Sub DisplayNode_InfoS_Worklist_KHNMC(asPath As String, asCnt As Integer)

    Dim xmlDoc          As New MSXML2.DOMDocument30
    Dim nodeBook        As IXMLDOMElement
    Dim nodeId          As IXMLDOMAttribute
    Dim xNode           As MSXML2.IXMLDOMNode
    Dim namedNodeMap    As IXMLDOMNamedNodeMap
    Dim Child_Node      As MSXML2.IXMLDOMNodeList
    Dim i, J            As Integer
    Dim intNodeLen      As Integer
    
On Error GoTo ErrXML:
    
    Set xmlDoc = New MSXML2.DOMDocument30
    
    xmlDoc.async = False
    xmlDoc.Load asPath
    xmlDoc
    
    If (xmlDoc.parseError.errorCode <> 0) Then
        Dim myErr
        Set myErr = xmlDoc.parseError
        
        MsgBox ("You have error " & myErr.reason)
    Else
        ReDim Preserve XmlSelectS.PID(asCnt)
        ReDim Preserve XmlSelectS.PATNM(asCnt)
        ReDim Preserve XmlSelectS.PRGSTNO(asCnt)
        ReDim Preserve XmlSelectS.BCNO(asCnt)
        ReDim Preserve XmlSelectS.TestCd(asCnt)
        
'        <pid>08558428</pid>
'        <kornm>�ڱ��</kornm>
'        <gsno>����corona-20-13636</gsno>
'        <bcno>2533416I0</bcno>
'        <ifcode>corona</ifcode>
        
        Set Child_Node = xmlDoc.childNodes
        For Each xNode In Child_Node
            If xNode.nodeType = NODE_ELEMENT Then
                For intNodeLen = 0 To xNode.childNodes.Length - 1
                    For i = 0 To xNode.childNodes.Item(intNodeLen).childNodes.Length - 1
                        'Debug.Print xNode.childNodes.Item(intNodeLen).childNodes.Item(i).baseName & ":" & xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue
                        Select Case UCase(xNode.childNodes.Item(intNodeLen).childNodes.Item(i).baseName)
                            Case "PID":         XmlSelectS.PID(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue                 'ȯ�ڹ�ȣ       [000132623]
                            Case "KORNM":       XmlSelectS.PATNM(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue               'ȯ�ڸ�         [�ڼ���]
                            Case "GSNO":        XmlSelectS.PRGSTNO(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue             '
                            Case "BCNO":        XmlSelectS.BCNO(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue                '��ü��ȣ       [3010700030]
                            Case "IFCODE":      XmlSelectS.TestCd(J) = xNode.childNodes.Item(intNodeLen).childNodes.Item(i).nodeTypedValue              '�˻��ڵ�       [VB6012A]
                        End Select
                    Next
                    J = J + 1
                Next
            End If
        Next
       
        Set Child_Node = Nothing
        
    End If

    Exit Sub
    
ErrXML:
    Exit Sub
    
End Sub

'-- �˻��� ���� ��������
Function GetSampleInfo_HCMS(ByVal asRow As Long, ByVal SPD As Object) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo        As Long
    Dim strWhere        As String
    
On Error GoTo DBErr
    
    GetSampleInfo_HCMS = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    strWhere = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
    
    With frmMain
        If .optStatus(3).Value = True Then
            strWhere = ""
        Else
            If .optStatus(0).Value = True Then
                strWhere = "32"
            End If
            If .optStatus(1).Value = True Then
                strWhere = "24"
            End If
            If .optStatus(2).Value = True Then
                strWhere = "12"
            End If
        End If
    End With
    
    SQL = ""
    SQL = SQL & "SELECT SUPT_DEPT_RCPN_DT                               AS HOSPDATE " & vbCrLf
    SQL = SQL & "     , PID                                             AS PID      " & vbCrLf
    SQL = SQL & "     , PT_NM                                           AS PNAME    " & vbCrLf
    SQL = SQL & "     , SPCM_RCPN_NO                                    AS BARCODE  " & vbCrLf
    SQL = SQL & "     , DEPT_HNM                                        AS DEPTNM   " & vbCrLf
    SQL = SQL & "     , IV_EXMN_CD                                      AS ITEM     " & vbCrLf
    SQL = SQL & "  FROM LIME.VW_READ_LSMACIF                                        " & vbCrLf
    SQL = SQL & " WHERE SPCM_RCPN_NO = '" & strBarcode & "'                         " & vbCrLf
    If strWhere <> "" Then
        SQL = SQL & "   AND STATUS = '" & strWhere & "'                             " & vbCrLf
    End If
    SQL = SQL & "   AND LAB_MAC_CODE = '" & gHOSP.PARTCD & "'                       " & vbCrLf
    SQL = SQL & "   AND IV_EXMN_CD   IN (" & gAllTestCd & ")                        " & vbCrLf
        
    Call SetSQLData("���ڵ���ȸ", SQL)
    
    '-- Record Count ������
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
                SetText SPD, Trim(RS.Fields("DEPTNM")) & "", asRow, colDEPT
                
                '��������
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '���������� ����
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '-- ȭ�鿡 ǥ��
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "��", asRow, intCol)
                        
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
    
    GetSampleInfo_HCMS = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_HCMS = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_GetSampleInfo_HCMS" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- �˻��� ���� ��������
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
        
    '-- SP ���
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
                
                '��������
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '���������� ����
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODENO")) & ""
                    .PID = Trim(RS.Fields("WORKNO")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                '-- ȭ�鿡 ǥ��
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEMCODE")) & Trim(RS.Fields("DCODE")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "��", asRow, intCol)
                        
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_GetSampleInfo_SCL" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
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
    SQL = SQL & "       (a.����� + a.����� + a.������)    AS HOSPDATE     " & vbCrLf
    SQL = SQL & "     , a.íƮ��ȣ                          AS CHARTNO      " & vbCrLf
    SQL = SQL & "     , c.�������                          AS STATE        " & vbCrLf
    SQL = SQL & "     , b.�����ڸ�                          AS PNAME        " & vbCrLf
    SQL = SQL & "     , b.�ֹε�Ϲ�ȣ                      AS PJUMIN       " & vbCrLf
    SQL = SQL & "     , (a.ó���ڵ� + a.�����ڵ�)           AS ITEM         " & vbCrLf
    SQL = SQL & "  From TB_�˻��׸� a, TB_�������� b, TB_����⺻ c         " & vbCrLf
    SQL = SQL & " Where a.íƮ��ȣ = '" & strChartNo & "'                   " & vbCrLf
    SQL = SQL & "   And a.ó���ȣ > 0                                      " & vbCrLf
    SQL = SQL & "   And c.������� IN ('1','5','6','7','8','9')             " & vbCrLf
'    SQL = SQL & "   And (a.ó���ڵ� + a.�����ڵ�) IN (" & gAllTestCd & ")   " & vbCrLf
    SQL = SQL & "   And (a.ó���ڵ� + '|' + a.�����ڵ�) IN (" & gAllTestCd & ")   " & vbCrLf
    SQL = SQL & "   And (a.�˻��� IS NULL OR a.�˻��� = '')             " & vbCrLf
    SQL = SQL & "   And a.�����    = c.�����                              " & vbCrLf
    SQL = SQL & "   And a.�����    = c.�����                              " & vbCrLf
    SQL = SQL & "   And a.������    = c.������                              " & vbCrLf
    SQL = SQL & "   And a.íƮ��ȣ  = c.íƮ��ȣ                            " & vbCrLf
    SQL = SQL & "   And a.íƮ��ȣ  = b.íƮ��ȣ                            " & vbCrLf
    SQL = SQL & "   And (a.�˻��� IS NULL OR a.�˻��� = '')             " & vbCrLf
    SQL = SQL & " Order By a.�����, a.�����, a.������, b.�����ڸ�         " & vbCrLf
        
    Call SetSQLData("���ڵ���ȸ", SQL)
    
    '-- Record Count ������
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
                
                '��������
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '���������� ����
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                'ȯ�� ����/����
                With mPatient
                    '.SEX = Trim(RS.Fields("SEX")) & ""
                    '.AGE = Trim(RS.Fields("AGE")) & ""
                End With
                
                '-- ȭ�鿡 ǥ��
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "��", asRow, intCol)
                        
                        '-- ó���ڵ�
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_GetSampleInfo_MEDICHART" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- �˻��� ���� ��������
Function GetSampleInfo_LABSPEAR(ByVal asRow As Long, ByVal SPD As Object) As Integer
    Dim strRegDate      As String
    Dim strBarcode      As String
    Dim strPatID        As String
    Dim strChartNo      As String
    Dim intCol          As Integer
    Dim intTestCnt      As Integer
    Dim lngRegNo            As Long
    
    Dim strRegNo            As String
    
On Error GoTo DBErr
    
    GetSampleInfo_LABSPEAR = -1
    
    intTestCnt = 0
    gPatOrdCd = ""
    
    strRegDate = Trim(GetText(SPD, asRow, colHOSPDATE))
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    strRegDate = Mid(strBarcode, 1, 8)
    strRegNo = Mid(strBarcode, 9)
    
    If strBarcode = "" Then
        Exit Function
    End If
              
    If strRegNo = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
        
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       CONVERT(NVARCHAR(50),M.��������,112)    AS HOSPDATE " & vbCrLf
    SQL = SQL & "     , M.������ȣ                              AS PID      " & vbCrLf
    SQL = SQL & "     , M.��Ʈ��ȣ                              AS CHARTNO  " & vbCrLf
    SQL = SQL & "     , M.����                                  AS PNAME    " & vbCrLf
    SQL = SQL & "     , M.����                                  AS SEX      " & vbCrLf
    SQL = SQL & "     , M.����                                  AS AGE      " & vbCrLf
    SQL = SQL & "     , M.�ŷ�ó��                              AS DEPT     " & vbCrLf
    SQL = SQL & "     , E.�˻��ڵ�                              AS ITEM     " & vbCrLf
    SQL = SQL & "  FROM VW_�˻����� M, VW_�˻��� R, VW_�˻��ڵ� E         " & vbCrLf
    SQL = SQL & " WHERE M.��������      = CONVERT(DATE, '" & strRegDate & "')" & vbCrLf
    SQL = SQL & "   AND M.������ȣ      = '" & strRegNo & "'                " & vbCrLf
    SQL = SQL & "   AND E.�к��ڵ�      = '" & gHOSP.PARTCD & "'            " & vbCrLf    'U2
    SQL = SQL & "   AND E.�˻��ڵ�      IN (" & gAllTestCd & ")             " & vbCrLf
    SQL = SQL & "   AND ISNULL(R.������, 'N') <> 'Y'                      " & vbCrLf
    SQL = SQL & "   AND (R.����� IS NULL OR R.����� = '')                 " & vbCrLf
    SQL = SQL & "   AND M.��������      = R.��������                        " & vbCrLf
    SQL = SQL & "   AND M.������ȣ      = R.������ȣ                        " & vbCrLf
    SQL = SQL & "   AND R.�˻��ڵ�      = E.�˻��ڵ�                        " & vbCrLf
   
    Call SetSQLData("���ڵ���ȸ", SQL)
    
    '-- Record Count ������
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
                '��������
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '���������� ����
                With mOrder
                    .BarNo = strBarcode 'Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                'ȯ�� ����/����
                With mPatient
                    .SEX = Trim(RS.Fields("SEX")) & ""
                    .AGE = Trim(RS.Fields("AGE")) & ""
                End With
                
                '-- ȭ�鿡 ǥ��
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "��", asRow, intCol)
                        
                        '-- ó���ڵ�
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_GetSampleInfo_LABSPEAR" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- �˻��� ���� ��������
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
    SQL = SQL & "   AND L.LABKEYNUM = D.DATKEYNUM                   " & vbCrLf      '-- ���̺���Ű��
    SQL = SQL & "   AND L.LABATTEND = D.DATATTEND                   " & vbCrLf      '-- ������ȣ
    SQL = SQL & "   AND L.LABATTEND = M.MANATTEND                   " & vbCrLf      '-- ������ȣ
    SQL = SQL & "   AND L.LABCHTNUM = D.DATCHTNUM                   " & vbCrLf      '-- íƮ��ȣ
    SQL = SQL & "   AND L.LABCHTNUM = M.MANCHTNUM                   " & vbCrLf      '-- íƮ��ȣ
    SQL = SQL & "   AND L.LABODRDTE = D.DATODRDTE                   " & vbCrLf      '-- ó������
    SQL = SQL & "   AND L.LABODRCOD IN (" & gAllTestCd & ")         " & vbCrLf
    SQL = SQL & "   AND (L.LABCANCEL = '' OR L.LABCANCEL IS NULL)   " & vbCrLf      '-- ��ҿ���
    SQL = SQL & "   AND (L.LABRESULT = '' OR L.LABRESULT IS NULL)   " & vbCrLf
    SQL = SQL & "   AND L.LABENDDEP < '3'                           " & vbCrLf      '-- ó������ (2:����, 3:����Է�)
        
        
    Call SetSQLData("���ڵ���ȸ", SQL)
    
    '-- Record Count ������
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
                    Case "A":   SetText SPD, "�ܷ�", asRow, colINOUT
                    Case "F":   SetText SPD, "�Կ�", asRow, colINOUT
                    Case Else:  SetText SPD, "", asRow, colINOUT
                End Select
                
                '��������
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '���������� ����
                With mOrder
                    .BarNo = Trim(RS.Fields("BARCODE")) & ""
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                'ȯ�� ����/����
                With mPatient
                    '-- �׽�Ʈ �� ����
                    'Call CalAgeSex(Trim(RS.Fields("JUMIN")) & "", Format(Now, "yyyy/mm/dd"))
                    
                    '.SEX = Trim(RS.Fields("SEX")) & ""
                    '.AGE = Trim(RS.Fields("AGE")) & ""
                End With
                
                '-- ȭ�鿡 ǥ��
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "��", asRow, intCol)
                        
                        '-- ó���ڵ�
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "_GetSampleInfo_BIT70" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- �˻��� ���� ��������
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
    strBarcode = Trim(GetText(SPD, asRow, colBARCODE))
    strPatID = Trim(GetText(SPD, asRow, colPID))
    strChartNo = Trim(GetText(SPD, asRow, colCHARTNO))
    
    If strBarcode = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
        
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       READING_YMD     AS HOSPDATE     " & vbCrLf
    SQL = SQL & "     , BCODE_NO        AS BARCODE      " & vbCrLf
    SQL = SQL & "     , PTNT_NO         AS PID          " & vbCrLf
    SQL = SQL & "     , PTNT_NM         AS PNAME        " & vbCrLf
    SQL = SQL & "     , AGE             AS AGE          " & vbCrLf
    SQL = SQL & "     , SEX             AS SEX          " & vbCrLf
    SQL = SQL & "     , IO_GB           AS INOUT        " & vbCrLf
    SQL = SQL & "     , ORD_CD          AS ITEM         " & vbCrLf
    SQL = SQL & "     , SP_CD           AS SPCCD        " & vbCrLf
    SQL = SQL & "  FROM LIS_INTERFACE1_V                " & vbCrLf
    SQL = SQL & " WHERE BCODE_NO = '" & strBarcode & "' " & vbCrLf
    SQL = SQL & "   AND STS_CD = '0'" & vbCrLf                      '0 ����, 1:�������
    SQL = SQL & "   AND ORD_CD IN (" & gAllTestCd & ") " & vbCrLf
    SQL = SQL & " ORDER BY ORD_CD " & vbCrLf
        
    Call SetSQLData("���ڵ���ȸ", SQL)
    
    '-- Record Count ������
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
                SetText SPD, IIf(Trim(RS.Fields("INOUT")) & "" = "10", "�Կ�", "�ܷ�"), asRow, colINOUT
                SetText SPD, Trim(RS.Fields("BARCODE")), asRow, colBARCODE
                SetText SPD, Trim(RS.Fields("PID")) & "", asRow, colPID
                SetText SPD, Trim(RS.Fields("PNAME")) & "", asRow, colPNAME
                SetText SPD, Trim(RS.Fields("SEX")) & "", asRow, colPSEX
                SetText SPD, Trim(RS.Fields("AGE")) & "", asRow, colPAGE
                
                '��������
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '���������� ����
                With mOrder
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                'ȯ�� ����/����
                With mPatient
                    .AGE = Trim(RS.Fields("AGE")) & ""
                    .SEX = Trim(RS.Fields("SEX")) & ""
                End With
                
                '-- ȭ�鿡 ǥ��
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "��", asRow, intCol)
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
    
    GetSampleInfo_MCC = 1
    
    Screen.MousePointer = 0
    
Exit Function

DBErr:
    GetSampleInfo_MCC = -1
    intTestCnt = 0
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "GetSampleInfo_MCC" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

'-- �˻��� ���� ��������
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
'    SQL = SQL & "   AND STS_CD = '0'" & vbCrLf                      '0 ����, 1:�������
'    SQL = SQL & "   AND ORD_CD IN (" & gAllTestCd & ") " & vbCrLf
'    SQL = SQL & " ORDER BY ORD_CD " & vbCrLf
        
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       C.JOBDATE                               AS HOSPDATE     " & vbCrLf
    SQL = SQL & "     , C.SPECNO                                AS BARCODE      " & vbCrLf
    SQL = SQL & "     , C.PTNO                                  AS CHARTNO      " & vbCrLf
    SQL = SQL & "     , C.JOBNO                                 AS PID          " & vbCrLf
    SQL = SQL & "     , DECODE(C.GBIO,'I','�Կ�','O','�ܷ�')    AS INOUT        " & vbCrLf
    SQL = SQL & "     , C.SNAME                                 AS PNAME        " & vbCrLf
    SQL = SQL & "     , C.SEX                                   AS SEX          " & vbCrLf
    SQL = SQL & "     , C.AGE                                   AS AGE          " & vbCrLf
    SQL = SQL & "     , A.MASTERCODE                            AS ORDERCODE    " & vbCrLf
    SQL = SQL & "     , A.SUBCODE                               AS ITEM         " & vbCrLf
    SQL = SQL & "  From TW_HSP_OCS.TWEXAM_RESULTC A                             " & vbCrLf
    SQL = SQL & "     , TW_HSP_OCS.TWEXAM_MASTER  B                             " & vbCrLf
    SQL = SQL & "     , TW_HSP_OCS.TWEXAM_SPECMST C                             " & vbCrLf
    SQL = SQL & " Where A.SPECNO = '" & strBarcode & "'                         " & vbCrLf
    'SQL = SQL & "   And B.EQUCODE1 = '" & gHOSP.MACHCD & "'                     " & vbCrLf '����ڵ�
    'SQL = SQL & "   AND A.MASTERCODE IN (" & gAllTestCd & ")                    " & vbCrLf
    SQL = SQL & "   AND A.SUBCODE IN (" & gAllTestCd & ")                       " & vbCrLf
    SQL = SQL & "   AND C.STATUS  <= '3'                                        " & vbCrLf '�˻����
    SQL = SQL & "   And C.SPECNO  = A.SPECNO                                    " & vbCrLf
    SQL = SQL & "   And A.MASTERCODE = B.MASTERCODE                             " & vbCrLf
    SQL = SQL & " ORDER BY C.JOBDATE, C.SPECNO                                  " & vbCrLf
        
        
    Call SetSQLData("���ڵ���ȸ", SQL)
    
    '-- Record Count ������
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
                
                '��������
                SetText SPD, CStr(intTestCnt), asRow, colOCNT
                                                                 
                '���������� ����
                With mOrder
                    .PID = Trim(RS.Fields("PID")) & ""
                    .PNAME = Trim(RS.Fields("PNAME")) & ""
                    .Count = CStr(intTestCnt)
                    .NoOrder = False
                End With
                
                'ȯ�� ����/����
                With mPatient
                    .AGE = Trim(RS.Fields("AGE")) & ""
                    .SEX = Trim(RS.Fields("SEX")) & ""
                End With
                
                '-- ȭ�鿡 ǥ��
                For intCol = colSTATE + 1 To .MaxCols
                    If Trim(RS.Fields("ITEM")) = gArrEQP(intCol - colSTATE, 2) Then
                        .Row = asRow
                        .Col = intCol
                        .BackColor = vbYellow
                        Call SetText(SPD, "��", asRow, intCol)
                        
                        '-- ó���ڵ�
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
    strErrMsg = strErrMsg & "��    ġ : " & gHOSP.MACHNM & "GetSampleInfo_TWIN" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "������ȣ : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "�������� : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

Function SetJudge(asResult As String, asEquipCode As String)

    Select Case gEMR
        Case "AMIS"                         '�ƹ̽�
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "EMEDI"                        '�̸޵�
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "BIT"                          '��Ʈ
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)

        Case "BIT70"                        '��Ʈ HIB70
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "EASYS"                        '������
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "EONM"                         '�̿¿�
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
            
        Case "GSEN"                         '����Ŀ�´����̼���(��íƮ)
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "HWASAN"                       'ȭ��
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "JAINCOM"                       '������
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "JWINFO"                       '�߿�����
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
            
        Case "KCHART"                       '�ٴ����Ʈ
                SetJudge = SetJudge_KCHART(asResult, asEquipCode)
        
        Case "KOMAIN"                       '�߿�����
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
            
        Case "KYU"                          '�Ǿ���б�����
                '��ũ����Ʈ ��ɾ���
                'SetJudge =  SetJudge_KYU(asResult,asEquipCode)
        Case "MEDICHART"                    '�޵�íƮ
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
            
        Case "MEDIIT"
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
            
        Case "MEDITOLISS"                    '
                SetJudge = SetJudge_MEDITOLISS(asResult, asEquipCode)
            
        Case "MSINFOTEC"                    'MS������
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
    Dim stryy, strmm, strdd, strDate  As String
    
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
    
'    SQL = SQL & "  L.����˻�ID AS R, " & vbCrLf
'    SQL = SQL & "  L.��������ID AS P, " & vbCrLf

    '���γ� ����ġ0~����ġ1,
    '���ο� ����ġ2~����ġ3,
    '�ҾƳ� ����ġ4~����ġ5,
    '�Ҿƿ� ����ġ6~����ġ7
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       A.ȯ�ڼ��� AS ����                                          " & vbCr
    SQL = SQL & "     , L.����ġ0, L.����ġ1, L.����ġ2, L.����ġ3                  " & vbCr
    SQL = SQL & "     , L.����ġ4, L.����ġ5, L.����ġ6, L.����ġ7                  " & vbCr
    SQL = SQL & "     , (L.ó���ڵ� + L.�����ڵ�) AS ITEM                           " & vbCr
    SQL = SQL & "  FROM             TB_����˻� L                                   " & vbCr
    SQL = SQL & "       INNER JOIN  TB_�������� J ON (L.��������ID = J.��������ID)  " & vbCr
    SQL = SQL & "       INNER JOIN  TB_�����Ϲ� A ON (J.��������   = A.��������     " & vbCr
    SQL = SQL & "                                AND  J.íƮ��ȣ   = A.íƮ��ȣ     " & vbCr
    SQL = SQL & "                                AND  J.�����ȣ   = A.�����ȣ)    " & vbCr
    SQL = SQL & "  Where L.��ü��ȣ = '" & mResult.BarNo & "'                       " & vbCr
    SQL = SQL & "    AND L.�˻���� < 5                                             " & vbCr
    SQL = SQL & "    AND (L.ó���ڵ� + L.�����ڵ�) = '" & sEquipCode & "'           " & vbCr
                                                                 

     Call SetSQLData("����ġ��ȸ", SQL)
     
     '-- Record Count ������
     AdoCn.CursorLocation = adUseClient
     Set RS1 = AdoCn.Execute(SQL, , 1)
     If Not RS1.EOF = True And Not RS1.BOF = True Then
         Do Until RS1.EOF
            strRefL = ""
            strRefH = ""
            If Trim(RS1.Fields("����")) & "" = "M" Then
                If Trim(RS1.Fields("����ġ0")) & "" <> "" Then
                    strRefL = Trim(RS1.Fields("����ġ0")) & ""
                    strRefH = Trim(RS1.Fields("����ġ1")) & ""
                End If
            Else
                If Trim(RS1.Fields("����")) & "" = "F" Then
                    If Trim(RS1.Fields("����ġ2")) & "" <> "" Then
                        strRefL = Trim(RS1.Fields("����ġ2")) & ""
                        strRefH = Trim(RS1.Fields("����ġ3")) & ""
                    Else
                        strRefL = Trim(RS1.Fields("����ġ0")) & ""
                        strRefH = Trim(RS1.Fields("����ġ1")) & ""
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
            Case "���Ծ���"
                    sResult = Trim(asResult)
            Case "����"
                    sResult = Trim(asResult)
            Case "����"
                    sResult = Trim(sResult)
            Case "����(����)"
                    sResult = asResult & "(" & Trim(sResult) & ")"
            Case "����(����)"
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

    With frmMain
        sExamDate = Format(Now, "yyyymmdd")
        If Trim(GetText(.spdOrder, asRow1, colSAVESEQ)) = "" Then
            Exit Function
        End If

        SQL = ""
        SQL = SQL & "DELETE FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO     = '" & gHOSP.MACHCD & "' " & vbCrLf
        SQL = SQL & "   AND EXAMDATE    = '" & Trim(GetText(.spdOrder, asRow1, colEXAMDATE)) & "' " & vbCrLf
        SQL = SQL & "   AND EXAMTIME    = '" & Trim(GetText(.spdOrder, asRow1, colEXAMTIME)) & "' " & vbCrLf
        SQL = SQL & "   AND SAVESEQ     = " & Trim(GetText(.spdOrder, asRow1, colSAVESEQ)) & vbCrLf
        SQL = SQL & "   AND HOSPDATE    = '" & Trim(GetText(.spdOrder, asRow1, colHOSPDATE)) & "' " & vbCrLf
        SQL = SQL & "   AND BARCODE     = '" & Trim(GetText(.spdOrder, asRow1, colBARCODE)) & "' " & vbCrLf
        SQL = SQL & "   AND EQUIPCODE   = '" & Trim(GetText(.spdResult, asRow2, colRCHANNEL)) & "'" & vbCrLf
        SQL = SQL & "   AND EXAMCODE    = '" & Trim(GetText(.spdResult, asRow2, colRTESTCD)) & "'" & vbCrLf

        If DBExec(AdoCn_Local, SQL) Then
            SQL = ""
            SQL = SQL & "INSERT INTO PATRESULT (" & vbCrLf
            SQL = SQL & "  EQUIPNO"                         '����ڵ�
            SQL = SQL & ", EXAMDATE"                        '�˻�����
            SQL = SQL & ", EXAMTIME"                        '�˻�ð�
            SQL = SQL & ", SAVESEQ"                         '�������(��¥��)
            SQL = SQL & ", HOSPDATE" & vbCrLf               '������������
            
            SQL = SQL & ", BARCODE"                         '��ü��ȣ
            SQL = SQL & ", PID"                             '���Ϲ�ȣ(������ȣ)
            SQL = SQL & ", CHARTNO"                         'íƮ��ȣ
            SQL = SQL & ", SPECIMEN"                        '��ü
            SQL = SQL & ", DEPT" & vbCrLf                   '�Ƿڰ�
            
            SQL = SQL & ", INOUT"                           '��/��
            SQL = SQL & ", ERYN"                            '���޿���
            SQL = SQL & ", RETESTYN"                        '��˿���
            SQL = SQL & ", PNAME"                           '�̸�
            SQL = SQL & ", PSEX" & vbCrLf                   '����(M,F)
            
            SQL = SQL & ", PAGE"                            '����
            SQL = SQL & ", EXAMUID"                         '�˻���ID
            SQL = SQL & ", DISKNO"                          'Rack
            SQL = SQL & ", POSNO"                           'Pos
            SQL = SQL & ", EQPSEQNO" & vbCrLf               '���˻��ȣ
            '============================================================
            
            SQL = SQL & ", SEQNO"                           '�˻��Ϸù�ȣ
            SQL = SQL & ", EQUIPCODE"                       '�˻�ä��
            SQL = SQL & ", ORDERCODE"                       '����ó���ڵ�
            SQL = SQL & ", EXAMCODE"                        '�����˻��ڵ�
            SQL = SQL & ", EXAMCODESUB" & vbCrLf            '�����˻��ڵ�(SUB)"
            
            SQL = SQL & ", EXAMNAME"                        '�˻��
            SQL = SQL & ", EQUIPRESULT"                     '�����"
            SQL = SQL & ", RESULT"                          '�Ҽ�������"
            SQL = SQL & ", PREVRESULT"                      '�������"
            SQL = SQL & ", REFJUDGE" & vbCrLf               '����(H,L)
            
            SQL = SQL & ", REFFLAG"                         'flag
            SQL = SQL & ", REFVALUE"                        '����ġ
            SQL = SQL & ", PANICVALUE"                      'Delta
            SQL = SQL & ", DELTAVALUE"                      'Panic
            SQL = SQL & ", SENDFLAG"                        '���۱���(0:������,1:����)"
            SQL = SQL & ", SENDDATE)" & vbCrLf               '��������
            
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
            
'            Call SetSQLData("��������", SQL)
            
            If Not DBExec(AdoCn_Local, SQL) Then
                Exit Function
            End If

        End If
    End With

End Function

'-- ���� �˻��� ��¥�� Max + 1 ��ȣ�� �����´�
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
