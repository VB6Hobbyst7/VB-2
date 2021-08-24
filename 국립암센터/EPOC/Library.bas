Attribute VB_Name = "Library"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Const CHART_HIDDEN = 1E+308

Public Type PatGen
    Age As String
    Sex As String
End Type
Public gPatGen As PatGen

Function SaveTransDataW(ByVal argSpcRow As Integer) As Integer
    Dim iRow            As Integer
    Dim lsID            As String
    Dim strDate         As String
    Dim strInNum        As String
    Dim strGumNum       As String
    Dim VallsID         As String
    Dim lsPid           As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strEqpCd        As String
    Dim strSubCD        As String
    Dim strRefVal       As String
    Dim strSpcCd        As String
    Dim strSex As String
    Dim strAge  As String
    Dim strORQN As String
    
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    
    Dim strReceNo   As String
    Dim strSeqNo   As String
    
    Dim tmpREF As String
    Dim strREF As String
    Dim GumEqpCd As String * 100
    
    Dim strExamDate As String
    Dim strHospDate As String
    
    Dim strKey1     As String
    Dim strKey2     As String
    Dim strSaveSeq  As String
    Dim strSubCodes As String
    Dim strChtNum   As String
    Dim strChannel As String
    Dim strReturn  As String
    Dim strRsltType As String
    Dim strUID      As String
    Dim strErr      As String
    
    Dim prm1 As New ADODB.Parameter
    Dim prm2 As New ADODB.Parameter
    Dim prm3 As New ADODB.Parameter
    Dim prm4 As New ADODB.Parameter
    Dim prm5 As New ADODB.Parameter
    Dim prm6 As New ADODB.Parameter
    Dim prm7 As New ADODB.Parameter
    Dim prm8 As New ADODB.Parameter
    Dim prm9 As New ADODB.Parameter
    
    Dim sParam  As String
    
    Dim strTemp As String
    
'On Error GoTo ErrHandle

    With frmInterface
        SaveTransDataW = -1
        
        lsID = Trim(GetText(.vasID, argSpcRow, colBARCODE))
        If InStr(lsID, "오더없음") > 0 Then
            Exit Function
        End If
        
        If lsID = "" Then
            Exit Function
        End If
        
        lsPid = Trim(GetText(.vasID, argSpcRow, colPID))
        strChtNum = Trim(GetText(.vasID, argSpcRow, colCHARTNO))
        strExamDate = Trim(GetText(.vasID, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.vasID, argSpcRow, colHOSPDATE))
        strSaveSeq = Trim(GetText(.vasID, argSpcRow, colSAVESEQ))
        
        '-- Local에서 환자별로 결과값 가져오기
        ClearSpread .vasTemp
        
              SQL = "SELECT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT,REFFLAG,PANICVALUE,DELTAVALUE,PSEX,SEQNO,PAGE,PID,DISKNO,POSNO,EXAMSUBCODE,INOUT,EXAMNAME,EXAMUID " & vbCrLf
        SQL = SQL & "  FROM PATRESULT " & vbCrLf
        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf                                           '장비코드
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'  " & vbCrLf                                      '검사일
        SQL = SQL & "   AND BARCODE = '" & lsID & "' " & vbCrLf       '바코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq       '저장번호
'        SQL = SQL & "   AND DISKNO = '" & Trim(GetText(.vasID, argSpcRow, colDISKNO)) & "' " & vbCrLf         'DISK 번호(진료검사ID)
'        SQL = SQL & "   AND POSNO = '" & Trim(GetText(.vasID, argSpcRow, colPOSNO)) & "' "                    'POS 번호(진료지원ID)
              
        res = GetDBSelectVas(gLocal, SQL, .vasTemp)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
                
        .vasTemp.MaxRows = .vasTemp.DataRowCnt + 1

        sResult = ""
        sResult1 = ""
        sResult2 = ""
        strKey1 = ""
        strKey2 = ""
        
        '서버로 결과값 저장하기
        For iRow = 1 To .vasTemp.DataRowCnt
            strChannel = Trim(GetText(.vasTemp, iRow, 1))
            strEqpCd = Trim(GetText(.vasTemp, iRow, 2))
            sResult1 = Trim(GetText(.vasTemp, iRow, 4)) '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, iRow, 3)) '결과(수정결과)
            strChannel = Trim(GetText(.vasTemp, iRow, 16))
            strSex = Trim(GetText(.vasTemp, iRow, 8))
            strAge = Trim(GetText(.vasTemp, iRow, 10))
            strORQN = Trim(GetText(.vasTemp, iRow, 14))
            strUID = Trim(GetText(.vasTemp, iRow, 17))
            
            '-- 장비결과적용
            'If .optSaveResult(0).Value = True Then
            '    sResult = sResult1
            'Else
                sResult = sResult2
            'End If
            
            '-- 서버저장
            If sResult <> "" Then
                sParam = "<Table>" & _
                        "<QID><![CDATA[PG_SRL.SLP91_P03]]></QID>" & _
                        "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                        "<USERID><![CDATA[SUA]]></USERID>" & _
                        "<EXECTYPE><![CDATA[NONQUERY]]></EXECTYPE>" & _
                        "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                        "<P0><![CDATA[" & lsID & "]]></P0>" & _
                        "<P1><![CDATA[" & strEqpCd & "]]></P1>" & _
                        "<P2><![CDATA[" & sResult & "]]></P2>" & _
                        "<P3><![CDATA[]]></P3>" & _
                        "<P4><![CDATA[" & gEquipCode & "]]></P4>" & _
                        "<P5><![CDATA[]]></P5>" & _
                        "<P6><![CDATA[]]></P6>" & _
                        "<P7><![CDATA[]]></P7>" & _
                        "<P8><![CDATA[]]></P8>" & _
                        "<P9><![CDATA[]]></P9>" & _
                        "</Table>"
                
                sParam = "<NewDataSet>" & sParam & "</NewDataSet>"

                Call Online_XML_Qry("PG_SRL.SLP91_P03", sParam)
                
                'Call SetSQLData("결과저장", sParam)
            End If
            
        Next iRow
        
        '-- 상태저장
        sParam = "<Table>" & _
                 "<QID><![CDATA[PG_SRL.SLP91_U07]]></QID>" & _
                 "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                 "<USERID><![CDATA[SUA]]></USERID>" & _
                 "<EXECTYPE><![CDATA[NONQUERY]]></EXECTYPE>" & _
                 "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                 "<P0><![CDATA[" & lsID & "]]></P0>" & _
                 "<P1><![CDATA[" & gIFUser & "]]></P1>" & _
                 "<P2><![CDATA[]]></P2>" & _
                 "<P3><![CDATA[]]></P3>" & _
                 "</Table>"

        sParam = "<NewDataSet>" & sParam & "</NewDataSet>"
        
        Call Online_XML_Qry("PG_SRL.SLP91_U07", sParam)
            
        'Call SetSQLData("상태저장", sParam)
            
        SaveTransDataW = 1
            
    End With

Exit Function

ErrHandle:
    SaveTransDataW = -1
    
End Function


'-- 검사자 정보 가져오기
Function GetSampleInfoW_NCC(ByVal asRow As Long) As Integer
    Dim sBarcode    As String
    Dim sChartNo    As String
    Dim GetOrderExamCode As String
    Dim intCol     As Integer
    Dim strTestCd   As String
    Dim pFrDt   As String
    Dim pToDt   As String
    Dim pFrNo   As String
    Dim pToNo   As String
    Dim strPatTableName As String
    Dim strDate As String
    
    GetSampleInfoW_NCC = -1
    
    sChartNo = Trim(GetText(frmInterface.vasID, asRow, colCHARTNO))
    strDate = Trim(GetText(frmInterface.vasID, asRow, colEXAMDATE))
    
    If sChartNo = "" Then
        Exit Function
    End If
    

    '-- 입원환자 담당의사 조회 : PG_SRL.SLP91_S28
    '==> 진료과,담당의 가져오기
    Call Online_XML(gXml_S28, sChartNo)
    
    '-- POCT 검사 발행 및 결과 입력: PG_SRL.SLP91_U06
    '==> 검체번호,접수번호 가져오기
    Call Online_XML(gXml_U06, sChartNo)
    
    sBarcode = gPat_Info_Select.BARCODE
    With frmInterface
        SetText .vasID, "1", asRow, colCHECKBOX
        SetText .vasID, sBarcode, asRow, colBARCODE
        SetText .vasID, gPat_Info_Select.RCPNO, asRow, colPID
        
        GetSampleInfoW_NCC = 1
        
    End With
    
    frmInterface.vasID.RowHeight(-1) = 12
    
End Function



Public Function SetSpace(asStr As String, asLen As Integer, Optional asPos As Integer = 1) As String
'asPos = 1 : Left 공백
'asPos = 2 : Right 공백 채우기
    Dim sTmp As String
    Dim i As Integer
    
    sTmp = ""
    If Len(asStr) >= asLen Then
        SetSpace = Left(asStr, asLen)
        Exit Function
    End If
    
    sTmp = asStr
    For i = 1 To asLen - Len(asStr)
        If asPos = 1 Then
            sTmp = " " & sTmp
        Else
            sTmp = sTmp & " "
        End If
    Next i
    
    SetSpace = sTmp
End Function

Public Function ChangeDateFormat(ByVal asStr As String, Optional argV As String = "/") As String
    If Len(asStr) = 10 Then
        ChangeDateFormat = Left(asStr, 4) & argV & Mid(asStr, 6, 2) & argV & Mid(asStr, 9, 2)
    ElseIf Len(asStr) = 8 Then
        ChangeDateFormat = Left(asStr, 4) & argV & Mid(asStr, 5, 2) & argV & Mid(asStr, 7, 2)
    End If
End Function

Public Sub InsertRow(ByVal vasTable As Object, ByVal argRow As Long)
'스프레드에 Row 추가
    vasTable.MaxRows = vasTable.MaxRows + 1
    vasTable.Row = argRow
    vasTable.Action = 7
End Sub

Public Sub DeleteRow(ByVal vasTable As Object, ByVal argRow1 As Integer, ByVal argRow2 As Integer)
'스프레드에 Row 삭제
    vasTable.Row = argRow1
    vasTable.Row2 = argRow2
    vasTable.Col = 1
    vasTable.Col2 = vasTable.MaxCols
    vasTable.BlockMode = True
    vasTable.Action = 5
    vasTable.BlockMode = False
End Sub

Public Sub Deletecol(ByVal vasTable As Object, ByVal argCol1 As Integer, ByVal argCol2 As Integer)
'스프레드에 Col 삭제
    vasTable.Row = 1 'argRow1
    vasTable.Row2 = vasTable.MaxRows ' argRow2
    vasTable.Col = argCol1  '1
    vasTable.Col2 = argCol2 'vasTable.MaxCols
    vasTable.BlockMode = True
    vasTable.Action = 6 '5
    vasTable.BlockMode = False
End Sub

Public Sub SelectFocus(ByRef argObj As Object)
'GetFocus 시 Object내의 Text가 전체 선택 되게 한다.
    argObj.SelStart = 0
    argObj.SelLength = Len(argObj.Text)
End Sub


Public Sub SaveData(ByVal argSQL As String, Optional argFlag As Integer = 0)
'argSQL의 내용을 파일로 저장
    Dim FilNum
        
    FilNum = FreeFile
    
    If Dir(App.Path & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.Path & "\Log")
    End If
    
    Open App.Path & "\Log\" & SeperatorCls(frmInterface.Text_Today.Text) & ".txt" For Append As FilNum
    Print #FilNum, Format(Time, "hh:nn:ss") & " " & argSQL
    Close FilNum
End Sub

Public Sub SaveQuery(argSQL As String, Optional argFlag As Integer = 0)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    
    FilNum = FreeFile
    
    If argFlag = 0 Then
        Open "c:\QueryErr.txt" For Output As FilNum
    Else
        Open "c:\QueryErr.txt" For Append As FilNum
    End If
    Print #FilNum, argSQL
    Close FilNum
End Sub

Public Function CR() As String
    CR = Chr(13) & Chr(10)
End Function

Public Function vasActiveCell(ByRef vasTable As Object, ByVal vasRow As Long, ByVal vasCol As Long) As Boolean
'특정 Cell 지정
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    vasTable.Action = 0
End Function

Public Function GetCurRow(ByRef vasTable As Object) As Integer
'현재 Active 된 Row 가져온다
    GetCurRow = vasTable.ActiveRow
End Function

Public Function GetCurCol(ByRef vasTable As Object) As Integer
'현재 Active 된 Col 가져온다
    GetCurCol = vasTable.ActiveCol
End Function

Public Sub ClearSpread(ByRef vasTable As Object, Optional argStartRow As Long = 1, Optional argStartCol As Long = 0)
'vsSpread의 내용을 Clear 한다.
    vasTable.Row = argStartRow
    vasTable.Col = argStartCol
    vasTable.Row2 = vasTable.DataRowCnt
    vasTable.Col2 = vasTable.DataColCnt
    vasTable.BlockMode = True
    vasTable.Action = 3
    vasTable.BlockMode = False
End Sub
Public Function SetText(ByRef vasTable As Object, ByVal SetStr As String, ByVal vasRow As Long, ByVal vasCol As Long) As Boolean
'vsSpread에 데이타 넣기
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    vasTable.Text = SetStr
End Function

Public Function GetText(ByRef vasTable As Object, ByVal vasRow As Long, ByVal vasCol As Long) As String
'vsSpread에서 데이타 가져오기
    If vasRow < 0 Or vasCol < 0 Then
        Exit Function
    End If
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    GetText = vasTable.Text
End Function

Public Function vasSort(ByRef vasTable As Object, ByVal key1 As Long, Optional key2 As Long = 0, Optional key3 As Long = 0, Optional key4 As Long = 0, Optional key5 As Long = 0) As Boolean
'정렬할 부분의 선택
    vasTable.Row = 0
    vasTable.Col = 0
    vasTable.Row2 = vasTable.DataRowCnt
    vasTable.Col2 = vasTable.DataColCnt
'정렬을 Row로 실시
    vasTable.SortBy = 2 'SS_SORT_BY_ROW
'정렬 키를 선택
    vasTable.SortKey(1) = key1
    vasTable.SortKeyOrder(1) = 1 'SS_SORT_ORDER_ASCENDING

    vasTable.SortKey(2) = key2
    If (key2 = 0) Then
        vasTable.SortKeyOrder(2) = 0
    Else
        vasTable.SortKeyOrder(2) = 1
    End If

    vasTable.SortKey(3) = key3
    If (key3 = 0) Then
        vasTable.SortKeyOrder(3) = 0
    Else
        vasTable.SortKeyOrder(3) = 1
    End If

    vasTable.SortKey(4) = key4
    If (key4 = 0) Then
        vasTable.SortKeyOrder(4) = 0
    Else
        vasTable.SortKeyOrder(4) = 1
    End If

    vasTable.SortKey(5) = key5
    If (key5 = 0) Then
        vasTable.SortKeyOrder(5) = 0
    Else
        vasTable.SortKeyOrder(5) = 1
    End If
'정렬
    vasTable.Action = 25 'SS_ACTION_SORT

    vasActiveCell vasTable, 1, 1
End Function

Public Function ScanCol(ByRef Obj As Object, ByVal SearchStr As String, _
                        ByVal colPos As Integer, Optional StartRow = 1) As Integer
'SpreadSheetd의 Col에 있는것과 같은 Text를 찾아낸다.
'Return : 같은 Text가 존재하면 그 Col,
'                     존재하지 않으면 -1 을 반환
    Dim i As Integer
    Dim ChkData As String

    For i = StartRow To Obj.DataRowCnt
        ChkData = Trim(GetText(Obj, i, colPos))
        If Trim(ChkData) = Trim(SearchStr) Then
            ScanCol = i
            Exit Function
        End If
    Next i
    
    ScanCol = -1
End Function

Public Sub DoSleep(Optional ByVal lMilliSec As Long = 0)
    'The DoSleep function allows other threads to have a time slice
    'and still keeps the main VB thread alive (since DPlay callbacks
    'run on separate threads outside of VB).
    Sleep lMilliSec
    DoEvents
End Sub

Public Function SeperatorCls(ByVal asStr As String) As String
'숫자외의 구분자를 모두 없앤다
    Dim i       As Integer
    Dim StrLen  As Integer
    Dim RtStr   As String
    
    RtStr = ""

    For i = 1 To Len(asStr)
        If IsNumeric(Mid(asStr, i, 1)) Then
            RtStr = RtStr & Mid(asStr, i, 1)
        End If
    Next i
    
    SeperatorCls = RtStr
End Function

Public Sub SetBackColor(asTable As vaSpread, ByVal asRow1 As Long, ByVal asRow2 As Long, ByVal asCol1 As Long, ByVal asCol2 As Long, asR As Variant, asG As Variant, asB As Variant)
    asTable.Row = asRow1
    asTable.Row2 = asRow2
    asTable.Col = asCol1
    asTable.Col2 = asCol2
    asTable.BlockMode = True
    asTable.BackColor = RGB(asR, asG, asB)
    asTable.BlockMode = False
End Sub

Public Sub SetForeColor(asTable As vaSpread, ByVal asRow1 As Long, ByVal asRow2 As Long, ByVal asCol1 As Long, ByVal asCol2 As Long, asR As Variant, asG As Variant, asB As Variant)
    asTable.Row = asRow1
    asTable.Row2 = asRow2
    asTable.Col = asCol1
    asTable.Col2 = asCol2
    asTable.BlockMode = True
    asTable.ForeColor = RGB(asR, asG, asB)
    asTable.BlockMode = False
End Sub

Public Function Data2Pict(sPrmData As String, sPrmPict As String) As String

    Dim i As Integer, iDataPos As Integer
    Dim iDataLen As Integer, iPictLen As Integer
    Dim sBufData As String, sPictStr As String, sChar As String

    iDataLen = Len(sPrmData)
    iPictLen = Len(sPrmPict)
    iDataPos = iDataLen
    sBufData = ""
    
    If iDataLen = 0 Or sPrmData = "0" Then
        If Right(sPrmPict, 1) = "0" Then
            Data2Pict = "0"
        Else
            Data2Pict = ""
        End If
        Exit Function
    End If

    For i = iPictLen To 1 Step -1
        sPictStr = ""

        Select Case Mid(sPrmPict, i, 1)
        Case "0", "9"
            sPictStr = Mid(sPrmData, iDataPos, 1)
            If Not IsNumeric(sPictStr) Then
                sPictStr = ""
                i = i + 1
            End If
            iDataPos = iDataPos - 1

        'Case ",", "."
        '    iDataPos = iDataPos - 1

        Case "X"
            sPictStr = Mid(sPrmData, iDataPos, 1)
            iDataPos = iDataPos - 1

        Case Else
            sPictStr = Mid(sPrmPict, i, 1)

        End Select

        sBufData = sPictStr & sBufData

        If iDataPos <= 0 Then
            Exit For
        End If
    Next

    If Left(LTrim(sPrmData), 1) = "-" Then
        sChar = Left(LTrim(sPrmPict), 1)
        Select Case sChar
        Case "-"
            If Left(LTrim(sBufData), 1) = "," Then
                sBufData = sChar & Mid(sBufData, 2)
            Else
                sBufData = sChar & sBufData
            End If

        End Select
    End If

    Data2Pict = sBufData

End Function

Public Function IsolateCode(argAll As String)
    Dim i As Integer
    Dim sCode, sName As String
    
    If argAll = "" Then
        gCode = ""
        gName = ""
        Exit Function
    End If
    
    sCode = ""
    sName = ""
    
    i = InStr(1, argAll, " ")
    
    If i = 0 Then
        gCode = Trim(argAll)
        gName = ""
    Else
        gCode = Trim(Left(argAll, i))
        gName = Trim(Mid(argAll, i))
    End If
End Function

Public Sub CalAgeSex(ByRef asPNRN As String, ByVal asCurDate As String)
    Dim sBirth As String
    Dim sStart As String
    
    gPatGen.Sex = ""
    gPatGen.Age = ""
    
    If Mid(asPNRN, 1, 1) = "_" Or Mid(asPNRN, 1, 1) = "" Then
        Exit Sub
    End If
        
    asPNRN = SeperatorCls(asPNRN)
    
    sStart = Trim(Mid(Trim(asPNRN), 9, 1))
    sBirth = Trim(Mid(Trim(asPNRN), 1, 8))
    
    Select Case sStart
        Case "1", "3", "5", "7"
            gPatGen.Sex = "M"
        Case "2", "4", "6", "8"
            gPatGen.Sex = "F"
    End Select

'    Select Case sStart
'        Case "1", "2"
'            sBirth = "19"
'        Case "3", "4"
'            sBirth = "20"
'        Case "7", "8"
'            sBirth = "18"
'        Case Else
'            sBirth = "19"
'    End Select
    
'    sBirth = ""
'    sBirth = sBirth & Mid(asPNRN, 1, 2) '& "/" & Mid(asPNRN, 3, 2) & "/" & Mid(asPNRN, 5, 2)
'    'If Mid(asPNRN, 3, 2) = "00" Then
'        sBirth = sBirth & "/01"
'    'Else
'    '    sBirth = sBirth & "/" & Mid(asPNRN, 3, 2)
'    'End If
'    'If Mid(asPNRN, 5, 2) = "00" Then
'        sBirth = sBirth & "/01"
'    'Else
'    '    sBirth = sBirth & "/" & Mid(asPNRN, 5, 2)
'    'End If
    
    gPatGen.Age = DateDiff("yyyy", sBirth, asCurDate) + 1
End Sub

Sub SetFont(asTable As vaSpread, ByVal asRow1 As Long, ByVal asRow2 As Long, ByVal asCol1 As Long, ByVal asCol2 As Long, asSize As Integer, asBold As Boolean)
    asTable.MaxRows = asTable.DataRowCnt
    
    asTable.Row = asRow1
    asTable.Row2 = asRow2
    asTable.Col = asCol1
    asTable.Col2 = asCol2
    asTable.BlockMode = True
    asTable.FontSize = asSize
    asTable.FontBold = asBold
    asTable.BlockMode = False
End Sub
