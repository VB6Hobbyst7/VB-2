Attribute VB_Name = "Library"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Type PatGen
    Age As String
    Sex As String
End Type

Public Const CHART_HIDDEN = 1E+308
Public gPatGen As PatGen

Public Sub KillProcess(PName As String)
    Dim pgm As String
    Dim wmi As Object
    Dim processes, process
    Dim sQuery As String

    pgm = PName
    Set wmi = GetObject("winmgmts:")
    sQuery = "select * from win32_process where name='" & pgm & "'"
    Set processes = wmi.execquery(sQuery)
    
    For Each process In processes
        process.Terminate
    Next

    Set wmi = Nothing
End Sub

Function ExecuteProcess(PName)
    Dim StrShell        As String
    Dim LngPID          As Long

On Error GoTo ExecuteNotePad_Error

    StrShell = "cmd /c " & App.Path & "\" & PName
    'StrShell = "cmd /c c:\1\" & PName
    LngPID = Shell(StrShell, vbHide)
    If LngPID = 0 Then
        'MsgBox StrShell & vbCrLf & vbCrLf & "½ÇÇà ¿À·ù", vbCritical, "¸Þ¸ðÀå ½ÇÇà ½ÇÆÐ"
    End If

    Exit Function

ExecuteNotePad_Error:
    MsgBox Err.Description, vbCritical, "CtrlAppInfo.ExecuteNotePad()"
    Exit Function
End Function


Function UTF8_Encode(ByRef sStr() As Byte) As String
    
    Dim ii As Long, sUTF8 As String, iChar As Long, iChar2 As Long
    
    For ii = 0 To UBound(sStr)
        iChar = sStr(ii)
        
        If iChar > 127 Then
            If Not iChar And 32 Then ' 2 chars
                iChar2 = sStr(ii + 1)
                sUTF8 = sUTF8 & ChrW$(((31 And iChar) * 64 + (63 And iChar2)))
                ii = ii + 1
            Else
                Dim iChar3 As Integer
                iChar2 = sStr(ii + 1)
                iChar3 = sStr(ii + 2)
                sUTF8 = sUTF8 & ChrW$(((iChar And 15) * 16 * 256) + ((iChar2 And 63) * 64) + (iChar3 And 63))
                ii = ii + 2
            End If
        Else
            sUTF8 = sUTF8 & Chr$(iChar)
        End If
    Next ii
    
    UTF8_Encode = sUTF8
    
End Function

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



Public Function SetSpace(asStr As String, asLen As Integer, Optional asPos As Integer = 1) As String
'asPos = 1 : Left °ø¹é
'asPos = 2 : Right °ø¹é Ã¤¿ì±â
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

Public Sub InsertRow(ByVal vasTable As Object, ByVal argRow As Integer)
'½ºÇÁ·¹µå¿¡ Row Ãß°¡
    vasTable.MaxRows = vasTable.MaxRows + 1
    vasTable.Row = argRow
    vasTable.Action = 7
End Sub

Public Sub DeleteRow(ByVal vasTable As Object, ByVal argRow1 As Integer, ByVal argRow2 As Integer)
'½ºÇÁ·¹µå¿¡ Row »èÁ¦
    vasTable.Row = argRow1
    vasTable.Row2 = argRow2
    vasTable.Col = 1
    vasTable.Col2 = vasTable.MaxCols
    vasTable.BlockMode = True
    vasTable.Action = 5
    vasTable.BlockMode = False
End Sub

Public Sub vasDeleteRow(ByVal vasTable As Object, argRow As Integer)
'Spread Row »èÁ¦
    vasTable.Row = argRow
    vasTable.Action = 5
End Sub

Public Sub SelectFocus(ByRef argObj As Object)
'GetFocus ½Ã Object³»ÀÇ Text°¡ ÀüÃ¼ ¼±ÅÃ µÇ°Ô ÇÑ´Ù.
    argObj.SelStart = 0
    argObj.SelLength = Len(argObj.Text)
End Sub

Public Sub SaveQuery(argSQL As String, Optional argFlag As Integer = 0)
'argSQLÀÇ ³»¿ëÀ» ÆÄÀÏ·Î ÀúÀå
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

Public Function vasActiveCell(ByRef vasTable As Object, ByVal vasRow As Integer, ByVal vasCol As Integer) As Boolean
'Æ¯Á¤ Cell ÁöÁ¤
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    vasTable.Action = 0
End Function

Public Function GetCurRow(ByRef vasTable As Object) As Integer
'ÇöÀç Active µÈ Row °¡Á®¿Â´Ù
    GetCurRow = vasTable.ActiveRow
End Function

Public Function GetCurCol(ByRef vasTable As Object) As Integer
'ÇöÀç Active µÈ Col °¡Á®¿Â´Ù
    GetCurCol = vasTable.ActiveCol
End Function

Public Sub DoSleep(Optional ByVal lMilliSec As Long = 0)
    'The DoSleep function allows other threads to have a time slice
    'and still keeps the main VB thread alive (since DPlay callbacks
    'run on separate threads outside of VB).
    Sleep lMilliSec
    DoEvents
End Sub

Public Sub ClearSpread(ByRef vasTable As Object, Optional argStartRow As Long = 1, Optional argStartCol As Long = 0)
'vsSpreadÀÇ ³»¿ëÀ» Clear ÇÑ´Ù.
    vasTable.Row = argStartRow
    vasTable.Col = argStartCol
    vasTable.Row2 = vasTable.DataRowCnt
    vasTable.Col2 = vasTable.DataColCnt
    vasTable.BlockMode = True
    vasTable.Action = 3
    vasTable.BlockMode = False
End Sub

Public Function SetText(ByRef vasTable As Object, ByVal SetStr As String, ByVal vasRow As Integer, ByVal vasCol As Integer) As Boolean
'vsSpread¿¡ µ¥ÀÌÅ¸ ³Ö±â
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    vasTable.Text = SetStr
End Function

Public Function GetText(ByRef vasTable As Object, ByVal vasRow As Integer, ByVal vasCol As Integer) As String
'vsSpread¿¡¼­ µ¥ÀÌÅ¸ °¡Á®¿À±â
    If vasRow < 0 Or vasCol < 0 Then
        Exit Function
    End If
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    GetText = vasTable.Text
End Function

Sub SetBackColor(asTable As vaSpread, ByVal asRow1 As Long, ByVal asRow2 As Long, asCol1 As Long, asCol2 As Long, asR As Variant, asG As Variant, asB As Variant)
    asTable.Row = asRow1
    asTable.Row2 = asRow2
    asTable.Col = asCol1
    asTable.Col2 = asCol2
    asTable.BlockMode = True
    asTable.BackColor = RGB(asR, asG, asB)
    asTable.BlockMode = False
End Sub

Sub SetForeColor(asTable As vaSpread, ByVal asRow1 As Long, ByVal asRow2 As Long, ByVal asCol1 As Long, ByVal asCol2 As Long, asR As Variant, asG As Variant, asB As Variant)
    asTable.Row = asRow1
    asTable.Row2 = asRow2
    asTable.Col = asCol1
    asTable.Col2 = asCol2
    asTable.BlockMode = True
    asTable.ForeColor = RGB(asR, asG, asB)
    asTable.BlockMode = False
End Sub

Public Function SeperatorCls(asStr As String) As String
'¼ýÀÚ¿ÜÀÇ ±¸ºÐÀÚ¸¦ ¸ðµÎ ¾ø¾Ø´Ù
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

Public Sub CalAgeSex(ByRef asPNRN As String, ByRef asCurDate As String)
    Dim sBirth As String
    Dim sStart As String
    
    gPatGen.Sex = ""
    gPatGen.Age = ""
    
    If Mid(asPNRN, 1, 1) = "_" Or Mid(asPNRN, 1, 1) = "" Then
        Exit Sub
    End If
        
    asPNRN = SeperatorCls(asPNRN)
    
    sStart = Trim(Mid(Trim(asPNRN), 7, 1))
    sBirth = ""
    
    Select Case sStart
        Case "1", "3", "5", "7"
            gPatGen.Sex = "M"
        Case "2", "4", "6", "8"
            gPatGen.Sex = "F"
    End Select

    Select Case sStart
        Case "1", "2"
            sBirth = "19"
        Case "3", "4"
            sBirth = "20"
        Case "7", "8"
            sBirth = "18"
        Case Else
            sBirth = "19"
    End Select
    
'    sBirth = ""
    sBirth = sBirth & Mid(asPNRN, 1, 2) '& "/" & Mid(asPNRN, 3, 2) & "/" & Mid(asPNRN, 5, 2)
    'If Mid(asPNRN, 3, 2) = "00" Then
        sBirth = sBirth & "/01"
    'Else
    '    sBirth = sBirth & "/" & Mid(asPNRN, 3, 2)
    'End If
    'If Mid(asPNRN, 5, 2) = "00" Then
        sBirth = sBirth & "/01"
    'Else
    '    sBirth = sBirth & "/" & Mid(asPNRN, 5, 2)
    'End If
    
    gPatGen.Age = DateDiff("yyyy", sBirth, asCurDate) + 1
End Sub


'Sub SetBackColor(asTable As vaSpread, ByVal asRow1 As Long, ByVal asRow2 As Long, asR As Variant, asG As Variant, asB As Variant)
'    asTable.Row = asRow1
'    asTable.Row2 = asRow2
'    asTable.Col = 1
'    asTable.Col2 = asTable.MaxCols
'    asTable.BlockMode = True
'    asTable.BackColor = RGB(asR, asG, asB)
'    asTable.BlockMode = False
'End Sub
'
'Sub SetForeColor(asTable As vaSpread, ByVal asRow1 As Long, ByVal asRow2 As Long, asR As Variant, asG As Variant, asB As Variant)
'    asTable.Row = asRow1
'    asTable.Row2 = asRow2
'    asTable.Col = 1
'    asTable.Col2 = asTable.MaxCols
'    asTable.BlockMode = True
'    asTable.ForeColor = RGB(asR, asG, asB)
'    asTable.BlockMode = False
'End Sub
Public Sub Save_Raw_Data(argSQL As String)
'argSQLÀÇ ³»¿ëÀ» ÆÄÀÏ·Î ÀúÀå
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
    
    If Dir(App.Path & "\Result", vbDirectory) <> "Result" Then
        MkDir (App.Path & "\Result")
    End If
    
    sFileName = gEquip & "_" & Format(Date, "yyyymmdd")
    
'    Open App.Path & "\Result\" & sFileName & ".txt" For Output As FilNum
    Open App.Path & "\Result\" & sFileName & ".txt" For Append As FilNum
    Print #FilNum, argSQL
    Close FilNum
End Sub

Public Sub Save_Raw_Data2(argSQL As String)
'argSQLÀÇ ³»¿ëÀ» ÆÄÀÏ·Î ÀúÀå
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
    
    If Dir(App.Path & "\Result", vbDirectory) <> "Result" Then
        MkDir (App.Path & "\Result")
    End If
    
    sFileName = gEquip & "_" & Format(Date, "yyyymmdd")
    
'    Open App.Path & "\Result\" & sFileName & ".txt" For Output As FilNum
    Open App.Path & "\Result\" & sFileName & "_Sub.txt" For Append As FilNum
    Print #FilNum, argSQL
    Close FilNum
End Sub

Public Function CheckDir(argDirName As String) As String
    CheckDir = "N"
    
    If Dir(App.Path & "\Result", vbDirectory) = "Result" Then
        CheckDir = "Y"
    End If
    
End Function

Public Sub Save_Trans_Data(argSQL As String)
'argSQLÀÇ ³»¿ëÀ» ÆÄÀÏ·Î ÀúÀå
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
    
    If Dir(App.Path & "\Trans", vbDirectory) <> "Trans" Then
        MkDir (App.Path & "\Trans")
    End If
    
    sFileName = gEquip & "_" & Format(Date, "yyyymmdd")
    
'    Open App.Path & "\Result\" & sFileName & ".txt" For Output As FilNum
    Open App.Path & "\Trans\" & sFileName & ".txt" For Append As FilNum
    Print #FilNum, argSQL
    Close FilNum
End Sub

Public Function vasSort(ByRef vasTable As Object, ByVal key1 As Integer, Optional key2 As Integer = 0, Optional key3 As Integer = 0, Optional key4 As Integer = 0, Optional key5 As Integer = 0) As Boolean
'Á¤·ÄÇÒ ºÎºÐÀÇ ¼±ÅÃ
    vasTable.Row = 0
    vasTable.Col = 0
    vasTable.Row2 = vasTable.DataRowCnt
    vasTable.Col2 = vasTable.DataColCnt
'Á¤·ÄÀ» Row·Î ½Ç½Ã
    vasTable.SortBy = 2 'SS_SORT_BY_ROW
'Á¤·Ä Å°¸¦ ¼±ÅÃ
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
'Á¤·Ä
    vasTable.Action = 25 'SS_ACTION_SORT

    vasActiveCell vasTable, 1, 1
End Function

Public Function Cut_KorEng(asData As String, asCount As Long) As String
    Dim sData As String
    Dim i, j As Integer
    Dim sTmp As String
    Dim sStrLen As Integer
    Dim Char As String
    Dim sCnt As Integer
    Dim X As Long
    
    Cut_KorEng = ""

    sData = asData
    
    sStrLen = LenB(sData)
    X = asCount
    
    sTmp = ""
    sCnt = 0
    For i = 1 To X
        If i > Len(sData) Then
            Exit For
        End If
        
        Char = Asc(Mid(sData, i, 1))
        If (Char >= 65 And Char <= 90) Or (Char >= 97 And Char <= 122) Or (Char >= 45 And Char <= 57) Or Char = 95 Or (Char >= 1 And Char <= 45) Then
        Else
            asCount = asCount - 1
        End If

    Next
    Cut_KorEng = Mid(asData, asCount)
    
    
End Function

Public Function Conv_Kor_Eng(ByVal asName As String) As String
    Dim sName As String
    Dim sEName As String
    Dim i As Integer
    
   
    On Error GoTo ErrName:
    
    sName = ""
    
    For i = 1 To Len(asName)
        sEName = hantoeng(Mid(asName, i, 1))
        sEName = UCase(Mid(sEName, 1, 1)) & Mid(sEName, 2)
        sName = sName & sEName
    
        If i = 1 Then
            If sName = "I" Then
                sName = "Lee"
            End If
        End If
    Next i
    
    Conv_Kor_Eng = Trim(sName)
    Exit Function
    
ErrName:
    Conv_Kor_Eng = ""
End Function

Public Function Conv_Kor_Eng_1(ByVal asName As String) As String
    Dim sName As String
    Dim i As Integer
    
    sName = ""
    
    For i = 1 To Len(asName)
        sName = sName & hantoeng(Mid(asName, i, 1)) & " "
    
        If i = 1 Then
            If sName = "I" Then
                sName = "lee"
            End If
        End If
    
    Next i
    
    Conv_Kor_Eng_1 = Trim(sName)
End Function

Public Function hantoeng(onehan As String) As String
       If Asc(onehan) >= 32 And Asc(onehan) < 135 Then
            hantoeng = onehan
            Exit Function
       End If
        
       Select Case onehan
       Case "°¡"
            hantoeng = "ga"
       Case "°¢"
            hantoeng = "gak"
       Case "°£"
            hantoeng = "gan"
       Case "°¥"
            hantoeng = "gal"
       Case "°¨"
            hantoeng = "gam"
       Case "°©"
            hantoeng = "gap"
       Case "°«"
            hantoeng = "gat"
       Case "°­"
            hantoeng = "gang"
       Case "°³"
            hantoeng = "gae"
       Case "°´"
            hantoeng = "gaek"
       Case "°Å"
            hantoeng = "geo"
       Case "°Ç"
            hantoeng = "geon"
       Case "°É"
            hantoeng = "geol"
       Case "°Ë"
            hantoeng = "geom"
       Case "°Ì"
            hantoeng = "geop"
       Case "°Ô"
            hantoeng = "ge"
       Case "°Ü"
            hantoeng = "gyeo"
       Case "°Ý"
            hantoeng = "gyeok"
       Case "°ß"
            hantoeng = "gyeon"
       Case "°á"
            hantoeng = "gyeol"
       Case "°â"
            hantoeng = "gyeom"
       Case "°ã"
            hantoeng = "gyeop"
       Case "°æ"
            hantoeng = "gyeong"
       Case "°è"
            hantoeng = "gye"
       Case "°í"
            hantoeng = "go"
       Case "°î"
            hantoeng = "gok"
       Case "°ï"
            hantoeng = "gon"
       Case "°ñ"
            hantoeng = "gol"
       Case "°÷"
            hantoeng = "got"
       Case "°ø"
            hantoeng = "gong"
       Case "°ù"
            hantoeng = "got"
       Case "°ú"
            hantoeng = "gwa"
       Case "°û"
            hantoeng = "gwak"
       Case "°ü"
            hantoeng = "gwan"
       Case "°ý"
            hantoeng = "gwal"
       Case "±¤"
            hantoeng = "gwang"
       Case "±¥"
            hantoeng = "gwae"
       Case "±«"
            hantoeng = "goe"
       Case "±²"
            hantoeng = "goeng"
       Case "±³"
            hantoeng = "gyo"
       Case "±¸"
            hantoeng = "gu"
       Case "±¹"
            hantoeng = "guk"
       Case "±º"
            hantoeng = "gun"
       Case "±¼"
            hantoeng = "gul"
       Case "±Â"
            hantoeng = "gut"
       Case "±Ã"
            hantoeng = "gung"
       Case "±Ç"
            hantoeng = "gwon"
       Case "±È"
            hantoeng = "gwol"
       Case "±Í"
            hantoeng = "gwi"
       Case "±Ô"
            hantoeng = "gyu"
       Case "±Õ"
            hantoeng = "gyun"
       Case "±Ö"
            hantoeng = "gyul"
       Case "±×"
            hantoeng = "geu"
       Case "±Ø"
            hantoeng = "geuk"
       Case "±Ù"
            hantoeng = "geun"
       Case "±Û"
            hantoeng = "geul"
       Case "±Û"
            hantoeng = "geul"
       Case "±Ý"
            hantoeng = "geum"
       Case "±Þ"
            hantoeng = "geup"
       Case "±à"
            hantoeng = "geung"
       Case "±â"
            hantoeng = "gi"
       Case "±ä"
            hantoeng = "gin"
       Case "±æ"
            hantoeng = "gil"
       Case "±è"
            hantoeng = "gim"
       Case "±î"
            hantoeng = "kka"
       Case "±ú"
            hantoeng = "kkae"
       Case "²¿"
            hantoeng = "kko"
       Case "²À"
            hantoeng = "kkok"
       Case "²É"
            hantoeng = "kkot"
       Case "²Ò"
            hantoeng = "kkoe"
       Case "²Ù"
            hantoeng = "kku"
       Case "²Þ"
            hantoeng = "kkum"
       Case "³¡"
            hantoeng = "kkeut"
       Case "³¢"
            hantoeng = "kki"
       Case "³ª"
            hantoeng = "na"
       Case "³«"
            hantoeng = "nak"
       Case "³­"
            hantoeng = "nan"
       Case "³¯"
            hantoeng = "nal"
       Case "³²"
            hantoeng = "nam"
       Case "³³"
            hantoeng = "nap"
       Case "³¶"
            hantoeng = "nang"
       Case "³»"
            hantoeng = "nae"
       Case "³Ã"
            hantoeng = "naeng"
       Case "³Ê"
            hantoeng = "neo"
       Case "³Î"
            hantoeng = "neol"
       Case "³×"
            hantoeng = "ne"
       Case "³à"
            hantoeng = "nyeo"
       Case "³á"
            hantoeng = "nyeok"
       Case "³â"
            hantoeng = "nyeon"
       Case "³ä"
            hantoeng = "nyeom"
       Case "³ç"
            hantoeng = "nyeong"
       Case "³ë"
            hantoeng = "no"
       Case "³ì"
            hantoeng = "nok"
       Case "³í"
            hantoeng = "non"
       Case "³î"
            hantoeng = "nol"
       Case "³ó"
            hantoeng = "nong"
       Case "³ú"
            hantoeng = "noe"
       Case "´©"
            hantoeng = "nu"
       Case "´«"
            hantoeng = "nun"
       Case "´­"
            hantoeng = "nul"
       Case "´À"
            hantoeng = "neu"
       Case "´Á"
            hantoeng = "neuk"
       Case "´Æ"
            hantoeng = "neum"
       Case "´É"
            hantoeng = "neung"
       Case "´Ì"
            hantoeng = "nui"
       Case "´Ï"
            hantoeng = "ni"
       Case "´Ð"
            hantoeng = "nik"
       Case "´Ñ"
            hantoeng = "nin"
       Case "´Ò"
            hantoeng = "nil"
       Case "´Ô"
            hantoeng = "nim"
       Case "´Ù"
            hantoeng = "da"
       Case "´Ü"
            hantoeng = "dan"
       Case "´Þ"
            hantoeng = "dal"
       Case "´ã"
            hantoeng = "dam"
       Case "´ä"
            hantoeng = "dap"
       Case "´ç"
            hantoeng = "dang"
       Case "´ë"
            hantoeng = "dae"
       Case "´ì"
            hantoeng = "daek"
       Case "´õ"
            hantoeng = "deo"
       Case "´ö"
            hantoeng = "deok"
       Case "µµ"
            hantoeng = "do"
       Case "µ¶"
            hantoeng = "dok"
       Case "µ·"
            hantoeng = "don"
       Case "µ¹"
            hantoeng = "dol"
       Case "µ¿"
            hantoeng = "dong"
       Case "µÅ"
            hantoeng = "dwae"
       Case "µÇ"
            hantoeng = "doe"
       Case "µÈ"
            hantoeng = "doen"
       Case "µÎ"
            hantoeng = "du"
       Case "µÏ"
            hantoeng = "duk"
       Case "µÐ"
            hantoeng = "dun"
       Case "µÚ"
            hantoeng = "dwi"
       Case "µå"
            hantoeng = "deu"
       Case "µæ"
            hantoeng = "deuk"
       Case "µé"
            hantoeng = "deul"
       Case "µî"
            hantoeng = "deung"
       Case "µð"
            hantoeng = "di"
       Case "µû"
            hantoeng = "tta"
       Case "¶¥"
            hantoeng = "ttang"
       Case "¶§"
            hantoeng = "ttae"
       Case "¶Ç"
            hantoeng = "tto"
       Case "¶Ñ"
            hantoeng = "ttu"
       Case "¶Ò"
            hantoeng = "ttuk"
       Case "¶ß"
            hantoeng = "tteu"
       Case "¶ì"
            hantoeng = "tti"
       Case "¶ó"
            hantoeng = "ra"
       Case "¶ô"
            hantoeng = "rak"
       Case "¶õ"
            hantoeng = "ran"
       Case "¶÷"
            hantoeng = "ram"
       Case "¶û"
            hantoeng = "rang"
       Case "·¡"
            hantoeng = "rae"
       Case "·©"
            hantoeng = "raeng"
       Case "·®"
            hantoeng = "ryang"
       Case "··"
            hantoeng = "reong"
       Case "·¹"
            hantoeng = "re"
       Case "·Á"
            hantoeng = "ryeo"
       Case "·Â"
            hantoeng = "ryeok"
       Case "·Ã"
            hantoeng = "ryeon"
       Case "·Ä"
            hantoeng = "ryeol"
       Case "·Å"
            hantoeng = "ryeom"
       Case "·Æ"
            hantoeng = "ryeop"
       Case "·É"
            hantoeng = "ryeong"
       Case "·Ê"
            hantoeng = "rye"
       Case "·Î"
            hantoeng = "ro"
       Case "·Ï"
            hantoeng = "rok"
       Case "·Ð"
            hantoeng = "ron"
       Case "·Õ"
            hantoeng = "rong"
       Case "·Ú"
            hantoeng = "roe"
       Case "·á"
            hantoeng = "ryo"
       Case "·æ"
            hantoeng = "ryong"
       Case "·ç"
            hantoeng = "ru"
       Case "·ù"
            hantoeng = "ryu"
       Case "·ú"
            hantoeng = "ryuk"
       Case "·û"
            hantoeng = "ryun"
       Case "·ü"
            hantoeng = "ryul"
       Case "¸¢"
            hantoeng = "ryung"
       Case "¸£"
            hantoeng = "reu"
       Case "¸¤"
            hantoeng = "reuk"
       Case "¸¥"
            hantoeng = "reun"
       Case "¸§"
            hantoeng = "reum"
       Case "¸ª"
            hantoeng = "reung"
       Case "¸®"
            hantoeng = "ri"
       Case "¸°"
            hantoeng = "rin"
       Case "¸²"
            hantoeng = "rim"
       Case "¸³"
            hantoeng = "rip"
       Case "¸¶"
            hantoeng = "ma"
       Case "¸·"
            hantoeng = "mak"
       Case "¸¸"
            hantoeng = "man"
       Case "¸»"
            hantoeng = "mal"
       Case "¸Á"
            hantoeng = "mang"
       Case "¸Å"
            hantoeng = "mae"
       Case "¸Æ"
            hantoeng = "maek"
       Case "¸Ç"
            hantoeng = "maen"
       Case "¸Í"
            hantoeng = "maeng"
       Case "¸Ó"
            hantoeng = "meo"
       Case "¸Ô"
            hantoeng = "meok"
       Case "¸Þ"
            hantoeng = "me"
       Case "¸ç"
            hantoeng = "myeo"
       Case "¸è"
            hantoeng = "myeok"
       Case "¸é"
            hantoeng = "myeon"
       Case "¸ê"
            hantoeng = "myeol"
       Case "¸í"
            hantoeng = "myeong"
       Case "¸ð"
            hantoeng = "mo"
       Case "¸ñ"
            hantoeng = "mok"
       Case "¸ô"
            hantoeng = "mol"
       Case "¸ø"
            hantoeng = "mot"
       Case "¸ù"
            hantoeng = "mong"
       Case "¸þ"
            hantoeng = "moe"
       Case "¹¦"
            hantoeng = "myo"
       Case "¹«"
            hantoeng = "mu"
       Case "¹¬"
            hantoeng = "muk"
       Case "¹®"
            hantoeng = "mun"
       Case "¹°"
            hantoeng = "mul"
       Case "¹Ç"
            hantoeng = "meu"
       Case "¹Ì"
            hantoeng = "mi"
       Case "¹Î"
            hantoeng = "min"
       Case "¹Ð"
            hantoeng = "mil"
       Case "¹Ù"
            hantoeng = "ba"
       Case "¹Ú"
            hantoeng = "bak"
       Case "¹Ý"
            hantoeng = "ban"
       Case "¹ß"
            hantoeng = "bal"
       Case "¹ä"
            hantoeng = "bap"
       Case "¹æ"
            hantoeng = "bang"
       Case "¹è"
            hantoeng = "bae"
       Case "¹é"
            hantoeng = "baek"
       Case "¹ì"
            hantoeng = "baem"
       Case "¹ö"
            hantoeng = "beo"
       Case "¹ø"
            hantoeng = "beon"
       Case "¹ú"
            hantoeng = "beol"
       Case "¹ü"
            hantoeng = "beom"
       Case "¹ý"
            hantoeng = "beop"
       Case "º­"
            hantoeng = "byeo"
       Case "º®"
            hantoeng = "byeok"
       Case "º¯"
            hantoeng = "byeon"
       Case "º°"
            hantoeng = "byeol"
       Case "º´"
            hantoeng = "byeong"
       Case "º¸"
            hantoeng = "bo"
       Case "º¹"
            hantoeng = "bok"
       Case "º»"
            hantoeng = "bon"
       Case "ºÀ"
            hantoeng = "bong"
       Case "ºÎ"
            hantoeng = "bu"
       Case "ºÏ"
            hantoeng = "buk"
       Case "ºÐ"
            hantoeng = "bun"
       Case "ºÒ"
            hantoeng = "bul"
       Case "ºØ"
            hantoeng = "bung"
       Case "ºñ"
            hantoeng = "bi"
       Case "ºó"
            hantoeng = "bin"
       Case "ºô"
            hantoeng = "bil"
       Case "ºö"
            hantoeng = "bim"
       Case "ºù"
            hantoeng = "bing"
       Case "ºü"
            hantoeng = "ppa"
       Case "»©"
            hantoeng = "ppae"
       Case "»µ"
            hantoeng = "ppeo"
       Case "»Ç"
            hantoeng = "ppo"
       Case "»Ñ"
            hantoeng = "ppu"
       Case "»Ú"
            hantoeng = "ppeu"
       Case "»ß"
            hantoeng = "ppi"
       Case "»ç"
            hantoeng = "sa"
       Case "»è"
            hantoeng = "sak"
       Case "»ê"
            hantoeng = "san"
       Case "»ì"
            hantoeng = "sal"
       Case "»ï"
            hantoeng = "sam"
       Case "»ð"
            hantoeng = "sap"
       Case "»ó"
            hantoeng = "sang"
       Case "»ô"
            hantoeng = "sat"
       Case "»õ"
            hantoeng = "sae"
       Case "»ö"
            hantoeng = "saek"
       Case "»ý"
            hantoeng = "saeng"
       Case "¼­"
            hantoeng = "seo"
       Case "¼®"
            hantoeng = "seok"
       Case "¼±"
            hantoeng = "seon"
       Case "¼³"
            hantoeng = "seol"
       Case "¼¶"
            hantoeng = "seom"
       Case "¼·"
            hantoeng = "seop"
       Case "¼º"
            hantoeng = "seong"
       Case "¼¼"
            hantoeng = "se"
       Case "¼Å"
            hantoeng = "syeo"
       Case "¼Ò"
            hantoeng = "so"
       Case "¼î"
            hantoeng = "syo"
       Case "¼Ó"
            hantoeng = "sok"
       Case "¼Õ"
            hantoeng = "son"
       Case "¼Ö"
            hantoeng = "sol"
       Case "¼Ø"
            hantoeng = "som"
       Case "¼Ú"
            hantoeng = "sot"
       Case "¼Û"
            hantoeng = "song"
       Case "¼â"
            hantoeng = "swae"
       Case "¼è"
            hantoeng = "soe"
       Case "¼ö"
            hantoeng = "su"
       Case "¼÷"
            hantoeng = "suk"
       Case "¼ø"
            hantoeng = "sun"
       Case "¼ú"
            hantoeng = "sul"
       Case "¼û"
            hantoeng = "sum"
       Case "¼þ"
            hantoeng = "sung"
       Case "½¬"
            hantoeng = "swi"
       Case "½º"
            hantoeng = "seu"
       Case "½½"
            hantoeng = "seul"
       Case "½¿"
            hantoeng = "seum"
       Case "½À"
            hantoeng = "seup"
       Case "½Â"
            hantoeng = "seung"
       Case "½Ã"
            hantoeng = "si"
       Case "½Ä"
            hantoeng = "sik"
       Case "½Å"
            hantoeng = "sin"
       Case "½Ç"
            hantoeng = "sil"
       Case "½É"
            hantoeng = "sim"
       Case "½Ê"
            hantoeng = "sip"
       Case "½Ì"
            hantoeng = "sing"
       Case "½Î"
            hantoeng = "ssa"
       Case "½Ö"
            hantoeng = "ssang"
       Case "½Ø"
            hantoeng = "ssae"
       Case "½î"
            hantoeng = "sso"
       Case "¾¦"
            hantoeng = "ssuk"
       Case "¾¾"
            hantoeng = "ssi"
       Case "¾Æ"
            hantoeng = "a"
       Case "¾Ç"
            hantoeng = "ak"
       Case "¾È"
            hantoeng = "an"
       Case "¾Ë"
            hantoeng = "al"
       Case "¾Ï"
            hantoeng = "am"
       Case "¾Ð"
            hantoeng = "ap"
       Case "¾Ó"
            hantoeng = "ang"
       Case "¾Õ"
            hantoeng = "ap"
       Case "¾Ö"
            hantoeng = "ae"
       Case "¾×"
            hantoeng = "aek"
       Case "¾Þ"
            hantoeng = "aeng"
       Case "¾ß"
            hantoeng = "ya"
       Case "¾à"
            hantoeng = "yak"
       Case "¾á"
            hantoeng = "yan"
       Case "¾ç"
            hantoeng = "yang"
       Case "¾î"
            hantoeng = "eo"
       Case "¾ï"
            hantoeng = "eok"
       Case "¾ð"
            hantoeng = "eon"
       Case "¾ó"
            hantoeng = "eol"
       Case "¾ö"
            hantoeng = "eom"
       Case "¾÷"
            hantoeng = "eop"
       Case "¿¡"
            hantoeng = "e"
       Case "¿¤"
            hantoeng = "el"
       Case "¿©"
            hantoeng = "yeo"
       Case "¿ª"
            hantoeng = "yeok"
       Case "¿¬"
            hantoeng = "yeon"
       Case "¿­"
            hantoeng = "yeol"
       Case "¿°"
            hantoeng = "yeom"
       Case "¿±"
            hantoeng = "yeop"
       Case "¿µ"
            hantoeng = "yeong"
       Case "¿¹"
            hantoeng = "ye"
       Case "¿À"
            hantoeng = "o"
       Case "¿Á"
            hantoeng = "ok"
       Case "¿Â"
            hantoeng = "on"
       Case "¿Ã"
            hantoeng = "ol"
       Case "¿È"
            hantoeng = "om"
       Case "¿Ë"
            hantoeng = "ong"
       Case "¿Í"
            hantoeng = "wa"
       Case "¿Ï"
            hantoeng = "wan"
       Case "¿Ð"
            hantoeng = "wal"
       Case "¿Õ"
            hantoeng = "wang"
       Case "¿Ö"
            hantoeng = "wae"
       Case "¿Ü"
            hantoeng = "oe"
       Case "¿Þ"
            hantoeng = "oen"
       Case "¿ä"
            hantoeng = "yo"
       Case "¿å"
            hantoeng = "yok"
       Case "¿ë"
            hantoeng = "yong"
       Case "¿ì"
            hantoeng = "u"
       Case "¿í"
            hantoeng = "uk"
       Case "¿î"
            hantoeng = "un"
       Case "¿ï"
            hantoeng = "ul"
       Case "¿ò"
            hantoeng = "um"
       Case "¿õ"
            hantoeng = "ung"
       Case "¿ö"
            hantoeng = "wo"
       Case "¿ø"
            hantoeng = "won"
       Case "¿ù"
            hantoeng = "wol"
       Case "À§"
            hantoeng = "wi"
       Case "À¯"
            hantoeng = "yu"
       Case "À°"
            hantoeng = "yuk"
       Case "À±"
            hantoeng = "yun"
       Case "À²"
            hantoeng = "yul"
       Case "À¶"
            hantoeng = "yung"
       Case "À·"
            hantoeng = "yut"
       Case "À¸"
            hantoeng = "eu"
       Case "Àº"
            hantoeng = "eun"
       Case "À»"
            hantoeng = "eul"
       Case "À½"
            hantoeng = "eum"
       Case "À¾"
            hantoeng = "eup"
       Case "ÀÀ"
            hantoeng = "eung"
       Case "ÀÇ"
            hantoeng = "ui"
       Case "ÀÌ"
            hantoeng = "I"
       Case "ÀÍ"
            hantoeng = "Ik"
       Case "ÀÎ"
            hantoeng = "In"
       Case "ÀÏ"
            hantoeng = "Il"
       Case "ÀÓ"
            hantoeng = "Im"
       Case "ÀÔ"
            hantoeng = "Ip"
       Case "À×"
            hantoeng = "Ing"
       Case "ÀÚ"
            hantoeng = "ja"
       Case "ÀÛ"
            hantoeng = "jak"
       Case "ÀÜ"
            hantoeng = "jan"
       Case "Àá"
            hantoeng = "jam"
       Case "Àâ"
            hantoeng = "jap"
       Case "Àå"
            hantoeng = "jang"
       Case "Àç"
            hantoeng = "jae"
       Case "Àï"
            hantoeng = "jaeng"
       Case "Àú"
            hantoeng = "jeo"
       Case "Àû"
            hantoeng = "jeok"
       Case "Àü"
            hantoeng = "jeon"
       Case "Àý"
            hantoeng = "jeol"
       Case "Á¡"
            hantoeng = "jeom"
       Case "Á¢"
            hantoeng = "jeop"
       Case "Á¤"
            hantoeng = "jeong"
       Case "Á¦"
            hantoeng = "je"
       Case "Á¶"
            hantoeng = "jo"
       Case "Á·"
            hantoeng = "jok"
       Case "Á¸"
            hantoeng = "jon"
       Case "Á¹"
            hantoeng = "jol"
       Case "Á¾"
            hantoeng = "jong"
       Case "ÁÂ"
            hantoeng = "jwa"
       Case "ÁË"
            hantoeng = "joe"
       Case "ÁÖ"
            hantoeng = "ju"
       Case "Á×"
            hantoeng = "juk"
       Case "ÁØ"
            hantoeng = "jun"
       Case "ÁÙ"
            hantoeng = "jul"
       Case "Áß"
            hantoeng = "jung"
       Case "Áã"
            hantoeng = "jwi"
       Case "Áî"
            hantoeng = "jeu"
       Case "Áï"
            hantoeng = "jeuk"
       Case "Áñ"
            hantoeng = "jeul"
       Case "Áò"
            hantoeng = "jeum"
       Case "Áó"
            hantoeng = "jeup"
       Case "Áõ"
            hantoeng = "jeung"
       Case "Áö"
            hantoeng = "ji"
       Case "Á÷"
            hantoeng = "jik"
       Case "Áø"
            hantoeng = "jin"
       Case "Áú"
            hantoeng = "jil"
       Case "Áü"
            hantoeng = "jim"
       Case "Áý"
            hantoeng = "jip"
       Case "Â¡"
            hantoeng = "jing"
       Case "Â¥"
            hantoeng = "jja"
       Case "Â°"
            hantoeng = "jjae"
       Case "ÂÉ"
            hantoeng = "jjo"
       Case "Âî"
            hantoeng = "jji"
       Case "Â÷"
            hantoeng = "cha"
       Case "Âø"
            hantoeng = "chak"
       Case "Âù"
            hantoeng = "chan"
       Case "Âû"
            hantoeng = "chal"
       Case "Âü"
            hantoeng = "cham"
       Case "Ã¢"
            hantoeng = "chang"
       Case "Ã¤"
            hantoeng = "chae"
       Case "Ã¥"
            hantoeng = "chaek"
       Case "Ã³"
            hantoeng = "cheo"
       Case "Ã´"
            hantoeng = "cheok"
       Case "Ãµ"
            hantoeng = "cheon"
       Case "Ã¶"
            hantoeng = "cheol"
       Case "Ã·"
            hantoeng = "cheom"
       Case "Ã¸"
            hantoeng = "cheop"
       Case "Ã»"
            hantoeng = "cheong"
       Case "Ã¼"
            hantoeng = "che"
       Case "ÃÊ"
            hantoeng = "cho"
       Case "ÃË"
            hantoeng = "chok"
       Case "ÃÌ"
            hantoeng = "chon"
       Case "ÃÑ"
            hantoeng = "chong"
       Case "ÃÖ"
            hantoeng = "choe"
       Case "Ãß"
            hantoeng = "chu"
       Case "Ãà"
            hantoeng = "chuk"
       Case "Ãá"
            hantoeng = "chun"
       Case "Ãâ"
            hantoeng = "chul"
       Case "Ãã"
            hantoeng = "chum"
       Case "Ãæ"
            hantoeng = "chung"
       Case "Ãø"
            hantoeng = "cheuk"
       Case "Ãø"
            hantoeng = "cheuk"
       Case "Ãþ"
            hantoeng = "cheung"
       Case "Ä¡"
            hantoeng = "chi"
       Case "Ä¢"
            hantoeng = "chik"
       Case "Ä£"
            hantoeng = "chin"
       Case "Ä¥"
            hantoeng = "chil"
       Case "Ä§"
            hantoeng = "chim"
       Case "Ä¨"
            hantoeng = "chip"
       Case "Äª"
            hantoeng = "ching"
       Case "ÄÚ"
            hantoeng = "ko"
       Case "Äè"
            hantoeng = "kwae"
       Case "Å©"
            hantoeng = "keu"
       Case "Å«"
            hantoeng = "keun"
       Case "Å°"
            hantoeng = "ki"
       Case "Å¸"
            hantoeng = "ta"
       Case "Å¹"
            hantoeng = "tak"
       Case "Åº"
            hantoeng = "tan"
       Case "Å»"
            hantoeng = "tal"
       Case "Å½"
            hantoeng = "tam"
       Case "Å¾"
            hantoeng = "tap"
       Case "ÅÁ"
            hantoeng = "tang"
       Case "ÅÂ"
            hantoeng = "tae"
       Case "ÅÃ"
            hantoeng = "taek"
       Case "ÅÊ"
            hantoeng = "taeng"
       Case "ÅÍ"
            hantoeng = "teo"
       Case "Å×"
            hantoeng = "te"
       Case "Åä"
            hantoeng = "to"
       Case "Åæ"
            hantoeng = "ton"
       Case "Åç"
            hantoeng = "tol"
       Case "Åë"
            hantoeng = "tong"
       Case "Åð"
            hantoeng = "toe"
       Case "Åõ"
            hantoeng = "tu"
       Case "Åü"
            hantoeng = "tung"
       Case "Æ¢"
            hantoeng = "twi"
       Case "Æ®"
            hantoeng = "teu"
       Case "Æ¯"
            hantoeng = "teuk"
       Case "Æ´"
            hantoeng = "teum"
       Case "Æ¼"
            hantoeng = "ti"
       Case "ÆÄ"
            hantoeng = "pa"
       Case "ÆÇ"
            hantoeng = "pan"
       Case "ÆÈ"
            hantoeng = "pal"
       Case "ÆÐ"
            hantoeng = "pae"
       Case "ÆØ"
            hantoeng = "paeng"
       Case "ÆÛ"
            hantoeng = "peo"
       Case "Æä"
            hantoeng = "pe"
       Case "Æì"
            hantoeng = "pyeo"
       Case "Æí"
            hantoeng = "pyeon"
       Case "Æï"
            hantoeng = "pyeom"
       Case "Æò"
            hantoeng = "pyeong"
       Case "Æó"
            hantoeng = "pye"
       Case "Æ÷"
            hantoeng = "po"
       Case "Æø"
            hantoeng = "pok"
       Case "Ç¥"
            hantoeng = "pyo"
       Case "Çª"
            hantoeng = "pu"
       Case "Ç°"
            hantoeng = "pum"
       Case "Ç³"
            hantoeng = "pung"
       Case "ÇÁ"
            hantoeng = "peu"
       Case "ÇÇ"
            hantoeng = "pi"
       Case "ÇÈ"
            hantoeng = "pik"
       Case "ÇÊ"
            hantoeng = "pil"
       Case "ÇÌ"
            hantoeng = "pip"
       Case "ÇÏ"
            hantoeng = "ha"
       Case "ÇÐ"
            hantoeng = "hak"
       Case "ÇÑ"
            hantoeng = "han"
       Case "ÇÒ"
            hantoeng = "hal"
       Case "ÇÔ"
            hantoeng = "ham"
       Case "ÇÕ"
            hantoeng = "hap"
       Case "Ç×"
            hantoeng = "hang"
       Case "ÇØ"
            hantoeng = "hae"
       Case "ÇÙ"
            hantoeng = "haek"
       Case "Çà"
            hantoeng = "haeng"
       Case "Çâ"
            hantoeng = "hyang"
       Case "Çã"
            hantoeng = "heo"
       Case "Çå"
            hantoeng = "heon"
       Case "Çè"
            hantoeng = "heom"
       Case "Çì"
            hantoeng = "he"
       Case "Çô"
            hantoeng = "hyeo"
       Case "Çõ"
            hantoeng = "hyeok"
       Case "Çö"
            hantoeng = "hyeon"
       Case "Ç÷"
            hantoeng = "hyeol"
       Case "Çø"
            hantoeng = "hyeom"
       Case "Çù"
            hantoeng = "hyeop"
       Case "Çü"
            hantoeng = "hyeong"
       Case "Çý"
            hantoeng = "hye"
       Case "È£"
            hantoeng = "ho"
       Case "È¤"
            hantoeng = "hok"
       Case "È¥"
            hantoeng = "hon"
       Case "È¦"
            hantoeng = "hol"
       Case "È©"
            hantoeng = "hop"
       Case "È«"
            hantoeng = "hong"
       Case "È­"
            hantoeng = "hwa"
       Case "È®"
            hantoeng = "hwak"
       Case "È¯"
            hantoeng = "hwan"
       Case "È°"
            hantoeng = "hwal"
       Case "È²"
            hantoeng = "hwang"
       Case "È³"
            hantoeng = "hwae"
       Case "È¶"
            hantoeng = "hwaet"
       Case "È¸"
            hantoeng = "hoe"
       Case "È¹"
            hantoeng = "hoek"
       Case "È¾"
            hantoeng = "hoeng"
       Case "È¿"
            hantoeng = "hyo"
       Case "ÈÄ"
            hantoeng = "hu"
       Case "ÈÆ"
            hantoeng = "hun"
       Case "ÈÍ"
            hantoeng = "hwon"
       Case "ÈÑ"
            hantoeng = "hwe"
       Case "ÈÖ"
            hantoeng = "hwi"
       Case "ÈÞ"
            hantoeng = "hyu"
       Case "Èá"
            hantoeng = "hyul"
       Case "Èä"
            hantoeng = "hyung"
       Case "Èå"
            hantoeng = "heu"
       Case "Èæ"
            hantoeng = "heuk"
       Case "Èç"
            hantoeng = "heun"
       Case "Èê"
            hantoeng = "heul"
       Case "Èì"
            hantoeng = "heum"
       Case "Èí"
            hantoeng = "heup"
       Case "Èï"
            hantoeng = "heung"
       Case "Èñ"
            hantoeng = "hui"
       Case "Èò"
            hantoeng = "huin"
       Case "È÷"
            hantoeng = "hi"
       Case "Èû"
            hantoeng = "him"
       
       Case " "     '2010.01.27 ÀÌ»óÀº Ãß°¡
            hantoeng = " "
       Case Else
            hantoeng = "??"
            
       End Select

End Function


