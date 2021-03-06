Attribute VB_Name = "Library"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Type PatGen
    Age As String
    Sex As String
End Type

Public Const CHART_HIDDEN = 1E+308
Public gPatGen As PatGen

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

Public Sub vasDeleteRow(ByVal vasTable As Object, argRow As Integer)
'Spread Row 삭제
    vasTable.Row = argRow
    vasTable.Action = 5
End Sub

Public Sub SelectFocus(ByRef argObj As Object)
'GetFocus 시 Object내의 Text가 전체 선택 되게 한다.
    argObj.SelStart = 0
    argObj.SelLength = Len(argObj.text)
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

Public Function vasActiveCell(ByRef vasTable As Object, ByVal vasRow As Integer, ByVal vasCol As Integer) As Boolean
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

Public Sub DoSleep(Optional ByVal lMilliSec As Long = 0)
    'The DoSleep function allows other threads to have a time slice
    'and still keeps the main VB thread alive (since DPlay callbacks
    'run on separate threads outside of VB).
    Sleep lMilliSec
    DoEvents
End Sub

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

Public Function SetText(ByRef vasTable As Object, ByVal SetStr As String, ByVal vasRow As Integer, ByVal vasCol As Integer) As Boolean
'vsSpread에 데이타 넣기
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    vasTable.text = SetStr
End Function

Public Function GetText(ByRef vasTable As Object, ByVal vasRow As Integer, ByVal vasCol As Integer) As String
'vsSpread에서 데이타 가져오기
    If vasRow < 0 Or vasCol < 0 Then
        Exit Function
    End If
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    GetText = vasTable.text
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

Sub SetForeColor(asTable As vaSpread, ByVal asRow1 As Long, ByVal asRow2 As Long, asR As Variant, asG As Variant, asB As Variant)
    asTable.Row = asRow1
    asTable.Row2 = asRow2
    asTable.Col = 1
    asTable.Col2 = asTable.MaxCols
    asTable.BlockMode = True
    asTable.ForeColor = RGB(asR, asG, asB)
    asTable.BlockMode = False
End Sub

Public Function SeperatorCls(asStr As String) As String
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
'argSQL의 내용을 파일로 저장
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

Public Sub Save_Trans_Data(argSQL As String)
'argSQL의 내용을 파일로 저장
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
       Case "가"
            hantoeng = "ga"
       Case "각"
            hantoeng = "gak"
       Case "간"
            hantoeng = "gan"
       Case "갈"
            hantoeng = "gal"
       Case "감"
            hantoeng = "gam"
       Case "갑"
            hantoeng = "gap"
       Case "갓"
            hantoeng = "gat"
       Case "강"
            hantoeng = "gang"
       Case "개"
            hantoeng = "gae"
       Case "객"
            hantoeng = "gaek"
       Case "거"
            hantoeng = "geo"
       Case "건"
            hantoeng = "geon"
       Case "걸"
            hantoeng = "geol"
       Case "검"
            hantoeng = "geom"
       Case "겁"
            hantoeng = "geop"
       Case "게"
            hantoeng = "ge"
       Case "겨"
            hantoeng = "gyeo"
       Case "격"
            hantoeng = "gyeok"
       Case "견"
            hantoeng = "gyeon"
       Case "결"
            hantoeng = "gyeol"
       Case "겸"
            hantoeng = "gyeom"
       Case "겹"
            hantoeng = "gyeop"
       Case "경"
            hantoeng = "gyeong"
       Case "계"
            hantoeng = "gye"
       Case "고"
            hantoeng = "go"
       Case "곡"
            hantoeng = "gok"
       Case "곤"
            hantoeng = "gon"
       Case "골"
            hantoeng = "gol"
       Case "곳"
            hantoeng = "got"
       Case "공"
            hantoeng = "gong"
       Case "곶"
            hantoeng = "got"
       Case "과"
            hantoeng = "gwa"
       Case "곽"
            hantoeng = "gwak"
       Case "관"
            hantoeng = "gwan"
       Case "괄"
            hantoeng = "gwal"
       Case "광"
            hantoeng = "gwang"
       Case "괘"
            hantoeng = "gwae"
       Case "괴"
            hantoeng = "goe"
       Case "굉"
            hantoeng = "goeng"
       Case "교"
            hantoeng = "gyo"
       Case "구"
            hantoeng = "gu"
       Case "국"
            hantoeng = "guk"
       Case "군"
            hantoeng = "gun"
       Case "굴"
            hantoeng = "gul"
       Case "굿"
            hantoeng = "gut"
       Case "궁"
            hantoeng = "gung"
       Case "권"
            hantoeng = "gwon"
       Case "궐"
            hantoeng = "gwol"
       Case "귀"
            hantoeng = "gwi"
       Case "규"
            hantoeng = "gyu"
       Case "균"
            hantoeng = "gyun"
       Case "귤"
            hantoeng = "gyul"
       Case "그"
            hantoeng = "geu"
       Case "극"
            hantoeng = "geuk"
       Case "근"
            hantoeng = "geun"
       Case "글"
            hantoeng = "geul"
       Case "글"
            hantoeng = "geul"
       Case "금"
            hantoeng = "geum"
       Case "급"
            hantoeng = "geup"
       Case "긍"
            hantoeng = "geung"
       Case "기"
            hantoeng = "gi"
       Case "긴"
            hantoeng = "gin"
       Case "길"
            hantoeng = "gil"
       Case "김"
            hantoeng = "gim"
       Case "까"
            hantoeng = "kka"
       Case "깨"
            hantoeng = "kkae"
       Case "꼬"
            hantoeng = "kko"
       Case "꼭"
            hantoeng = "kkok"
       Case "꽃"
            hantoeng = "kkot"
       Case "꾀"
            hantoeng = "kkoe"
       Case "꾸"
            hantoeng = "kku"
       Case "꿈"
            hantoeng = "kkum"
       Case "끝"
            hantoeng = "kkeut"
       Case "끼"
            hantoeng = "kki"
       Case "나"
            hantoeng = "na"
       Case "낙"
            hantoeng = "nak"
       Case "난"
            hantoeng = "nan"
       Case "날"
            hantoeng = "nal"
       Case "남"
            hantoeng = "nam"
       Case "납"
            hantoeng = "nap"
       Case "낭"
            hantoeng = "nang"
       Case "내"
            hantoeng = "nae"
       Case "냉"
            hantoeng = "naeng"
       Case "너"
            hantoeng = "neo"
       Case "널"
            hantoeng = "neol"
       Case "네"
            hantoeng = "ne"
       Case "녀"
            hantoeng = "nyeo"
       Case "녁"
            hantoeng = "nyeok"
       Case "년"
            hantoeng = "nyeon"
       Case "념"
            hantoeng = "nyeom"
       Case "녕"
            hantoeng = "nyeong"
       Case "노"
            hantoeng = "no"
       Case "녹"
            hantoeng = "nok"
       Case "논"
            hantoeng = "non"
       Case "놀"
            hantoeng = "nol"
       Case "농"
            hantoeng = "nong"
       Case "뇌"
            hantoeng = "noe"
       Case "누"
            hantoeng = "nu"
       Case "눈"
            hantoeng = "nun"
       Case "눌"
            hantoeng = "nul"
       Case "느"
            hantoeng = "neu"
       Case "늑"
            hantoeng = "neuk"
       Case "늠"
            hantoeng = "neum"
       Case "능"
            hantoeng = "neung"
       Case "늬"
            hantoeng = "nui"
       Case "니"
            hantoeng = "ni"
       Case "닉"
            hantoeng = "nik"
       Case "닌"
            hantoeng = "nin"
       Case "닐"
            hantoeng = "nil"
       Case "님"
            hantoeng = "nim"
       Case "다"
            hantoeng = "da"
       Case "단"
            hantoeng = "dan"
       Case "달"
            hantoeng = "dal"
       Case "담"
            hantoeng = "dam"
       Case "답"
            hantoeng = "dap"
       Case "당"
            hantoeng = "dang"
       Case "대"
            hantoeng = "dae"
       Case "댁"
            hantoeng = "daek"
       Case "더"
            hantoeng = "deo"
       Case "덕"
            hantoeng = "deok"
       Case "도"
            hantoeng = "do"
       Case "독"
            hantoeng = "dok"
       Case "돈"
            hantoeng = "don"
       Case "돌"
            hantoeng = "dol"
       Case "동"
            hantoeng = "dong"
       Case "돼"
            hantoeng = "dwae"
       Case "되"
            hantoeng = "doe"
       Case "된"
            hantoeng = "doen"
       Case "두"
            hantoeng = "du"
       Case "둑"
            hantoeng = "duk"
       Case "둔"
            hantoeng = "dun"
       Case "뒤"
            hantoeng = "dwi"
       Case "드"
            hantoeng = "deu"
       Case "득"
            hantoeng = "deuk"
       Case "들"
            hantoeng = "deul"
       Case "등"
            hantoeng = "deung"
       Case "디"
            hantoeng = "di"
       Case "따"
            hantoeng = "tta"
       Case "땅"
            hantoeng = "ttang"
       Case "때"
            hantoeng = "ttae"
       Case "또"
            hantoeng = "tto"
       Case "뚜"
            hantoeng = "ttu"
       Case "뚝"
            hantoeng = "ttuk"
       Case "뜨"
            hantoeng = "tteu"
       Case "띠"
            hantoeng = "tti"
       Case "라"
            hantoeng = "ra"
       Case "락"
            hantoeng = "rak"
       Case "란"
            hantoeng = "ran"
       Case "람"
            hantoeng = "ram"
       Case "랑"
            hantoeng = "rang"
       Case "래"
            hantoeng = "rae"
       Case "랭"
            hantoeng = "raeng"
       Case "량"
            hantoeng = "ryang"
       Case "렁"
            hantoeng = "reong"
       Case "레"
            hantoeng = "re"
       Case "려"
            hantoeng = "ryeo"
       Case "력"
            hantoeng = "ryeok"
       Case "련"
            hantoeng = "ryeon"
       Case "렬"
            hantoeng = "ryeol"
       Case "렴"
            hantoeng = "ryeom"
       Case "렵"
            hantoeng = "ryeop"
       Case "령"
            hantoeng = "ryeong"
       Case "례"
            hantoeng = "rye"
       Case "로"
            hantoeng = "ro"
       Case "록"
            hantoeng = "rok"
       Case "론"
            hantoeng = "ron"
       Case "롱"
            hantoeng = "rong"
       Case "뢰"
            hantoeng = "roe"
       Case "료"
            hantoeng = "ryo"
       Case "룡"
            hantoeng = "ryong"
       Case "루"
            hantoeng = "ru"
       Case "류"
            hantoeng = "ryu"
       Case "륙"
            hantoeng = "ryuk"
       Case "륜"
            hantoeng = "ryun"
       Case "률"
            hantoeng = "ryul"
       Case "륭"
            hantoeng = "ryung"
       Case "르"
            hantoeng = "reu"
       Case "륵"
            hantoeng = "reuk"
       Case "른"
            hantoeng = "reun"
       Case "름"
            hantoeng = "reum"
       Case "릉"
            hantoeng = "reung"
       Case "리"
            hantoeng = "ri"
       Case "린"
            hantoeng = "rin"
       Case "림"
            hantoeng = "rim"
       Case "립"
            hantoeng = "rip"
       Case "마"
            hantoeng = "ma"
       Case "막"
            hantoeng = "mak"
       Case "만"
            hantoeng = "man"
       Case "말"
            hantoeng = "mal"
       Case "망"
            hantoeng = "mang"
       Case "매"
            hantoeng = "mae"
       Case "맥"
            hantoeng = "maek"
       Case "맨"
            hantoeng = "maen"
       Case "맹"
            hantoeng = "maeng"
       Case "머"
            hantoeng = "meo"
       Case "먹"
            hantoeng = "meok"
       Case "메"
            hantoeng = "me"
       Case "며"
            hantoeng = "myeo"
       Case "멱"
            hantoeng = "myeok"
       Case "면"
            hantoeng = "myeon"
       Case "멸"
            hantoeng = "myeol"
       Case "명"
            hantoeng = "myeong"
       Case "모"
            hantoeng = "mo"
       Case "목"
            hantoeng = "mok"
       Case "몰"
            hantoeng = "mol"
       Case "못"
            hantoeng = "mot"
       Case "몽"
            hantoeng = "mong"
       Case "뫼"
            hantoeng = "moe"
       Case "묘"
            hantoeng = "myo"
       Case "무"
            hantoeng = "mu"
       Case "묵"
            hantoeng = "muk"
       Case "문"
            hantoeng = "mun"
       Case "물"
            hantoeng = "mul"
       Case "므"
            hantoeng = "meu"
       Case "미"
            hantoeng = "mi"
       Case "민"
            hantoeng = "min"
       Case "밀"
            hantoeng = "mil"
       Case "바"
            hantoeng = "ba"
       Case "박"
            hantoeng = "bak"
       Case "반"
            hantoeng = "ban"
       Case "발"
            hantoeng = "bal"
       Case "밥"
            hantoeng = "bap"
       Case "방"
            hantoeng = "bang"
       Case "배"
            hantoeng = "bae"
       Case "백"
            hantoeng = "baek"
       Case "뱀"
            hantoeng = "baem"
       Case "버"
            hantoeng = "beo"
       Case "번"
            hantoeng = "beon"
       Case "벌"
            hantoeng = "beol"
       Case "범"
            hantoeng = "beom"
       Case "법"
            hantoeng = "beop"
       Case "벼"
            hantoeng = "byeo"
       Case "벽"
            hantoeng = "byeok"
       Case "변"
            hantoeng = "byeon"
       Case "별"
            hantoeng = "byeol"
       Case "병"
            hantoeng = "byeong"
       Case "보"
            hantoeng = "bo"
       Case "복"
            hantoeng = "bok"
       Case "본"
            hantoeng = "bon"
       Case "봉"
            hantoeng = "bong"
       Case "부"
            hantoeng = "bu"
       Case "북"
            hantoeng = "buk"
       Case "분"
            hantoeng = "bun"
       Case "불"
            hantoeng = "bul"
       Case "붕"
            hantoeng = "bung"
       Case "비"
            hantoeng = "bi"
       Case "빈"
            hantoeng = "bin"
       Case "빌"
            hantoeng = "bil"
       Case "빔"
            hantoeng = "bim"
       Case "빙"
            hantoeng = "bing"
       Case "빠"
            hantoeng = "ppa"
       Case "빼"
            hantoeng = "ppae"
       Case "뻐"
            hantoeng = "ppeo"
       Case "뽀"
            hantoeng = "ppo"
       Case "뿌"
            hantoeng = "ppu"
       Case "쁘"
            hantoeng = "ppeu"
       Case "삐"
            hantoeng = "ppi"
       Case "사"
            hantoeng = "sa"
       Case "삭"
            hantoeng = "sak"
       Case "산"
            hantoeng = "san"
       Case "살"
            hantoeng = "sal"
       Case "삼"
            hantoeng = "sam"
       Case "삽"
            hantoeng = "sap"
       Case "상"
            hantoeng = "sang"
       Case "샅"
            hantoeng = "sat"
       Case "새"
            hantoeng = "sae"
       Case "색"
            hantoeng = "saek"
       Case "생"
            hantoeng = "saeng"
       Case "서"
            hantoeng = "seo"
       Case "석"
            hantoeng = "seok"
       Case "선"
            hantoeng = "seon"
       Case "설"
            hantoeng = "seol"
       Case "섬"
            hantoeng = "seom"
       Case "섭"
            hantoeng = "seop"
       Case "성"
            hantoeng = "seong"
       Case "세"
            hantoeng = "se"
       Case "셔"
            hantoeng = "syeo"
       Case "소"
            hantoeng = "so"
       Case "쇼"
            hantoeng = "syo"
       Case "속"
            hantoeng = "sok"
       Case "손"
            hantoeng = "son"
       Case "솔"
            hantoeng = "sol"
       Case "솜"
            hantoeng = "som"
       Case "솟"
            hantoeng = "sot"
       Case "송"
            hantoeng = "song"
       Case "쇄"
            hantoeng = "swae"
       Case "쇠"
            hantoeng = "soe"
       Case "수"
            hantoeng = "su"
       Case "숙"
            hantoeng = "suk"
       Case "순"
            hantoeng = "sun"
       Case "술"
            hantoeng = "sul"
       Case "숨"
            hantoeng = "sum"
       Case "숭"
            hantoeng = "sung"
       Case "쉬"
            hantoeng = "swi"
       Case "스"
            hantoeng = "seu"
       Case "슬"
            hantoeng = "seul"
       Case "슴"
            hantoeng = "seum"
       Case "습"
            hantoeng = "seup"
       Case "승"
            hantoeng = "seung"
       Case "시"
            hantoeng = "si"
       Case "식"
            hantoeng = "sik"
       Case "신"
            hantoeng = "sin"
       Case "실"
            hantoeng = "sil"
       Case "심"
            hantoeng = "sim"
       Case "십"
            hantoeng = "sip"
       Case "싱"
            hantoeng = "sing"
       Case "싸"
            hantoeng = "ssa"
       Case "쌍"
            hantoeng = "ssang"
       Case "쌔"
            hantoeng = "ssae"
       Case "쏘"
            hantoeng = "sso"
       Case "쑥"
            hantoeng = "ssuk"
       Case "씨"
            hantoeng = "ssi"
       Case "아"
            hantoeng = "a"
       Case "악"
            hantoeng = "ak"
       Case "안"
            hantoeng = "an"
       Case "알"
            hantoeng = "al"
       Case "암"
            hantoeng = "am"
       Case "압"
            hantoeng = "ap"
       Case "앙"
            hantoeng = "ang"
       Case "앞"
            hantoeng = "ap"
       Case "애"
            hantoeng = "ae"
       Case "액"
            hantoeng = "aek"
       Case "앵"
            hantoeng = "aeng"
       Case "야"
            hantoeng = "ya"
       Case "약"
            hantoeng = "yak"
       Case "얀"
            hantoeng = "yan"
       Case "양"
            hantoeng = "yang"
       Case "어"
            hantoeng = "eo"
       Case "억"
            hantoeng = "eok"
       Case "언"
            hantoeng = "eon"
       Case "얼"
            hantoeng = "eol"
       Case "엄"
            hantoeng = "eom"
       Case "업"
            hantoeng = "eop"
       Case "에"
            hantoeng = "e"
       Case "엘"
            hantoeng = "el"
       Case "여"
            hantoeng = "yeo"
       Case "역"
            hantoeng = "yeok"
       Case "연"
            hantoeng = "yeon"
       Case "열"
            hantoeng = "yeol"
       Case "염"
            hantoeng = "yeom"
       Case "엽"
            hantoeng = "yeop"
       Case "영"
            hantoeng = "yeong"
       Case "예"
            hantoeng = "ye"
       Case "오"
            hantoeng = "o"
       Case "옥"
            hantoeng = "ok"
       Case "온"
            hantoeng = "on"
       Case "올"
            hantoeng = "ol"
       Case "옴"
            hantoeng = "om"
       Case "옹"
            hantoeng = "ong"
       Case "와"
            hantoeng = "wa"
       Case "완"
            hantoeng = "wan"
       Case "왈"
            hantoeng = "wal"
       Case "왕"
            hantoeng = "wang"
       Case "왜"
            hantoeng = "wae"
       Case "외"
            hantoeng = "oe"
       Case "왼"
            hantoeng = "oen"
       Case "요"
            hantoeng = "yo"
       Case "욕"
            hantoeng = "yok"
       Case "용"
            hantoeng = "yong"
       Case "우"
            hantoeng = "u"
       Case "욱"
            hantoeng = "uk"
       Case "운"
            hantoeng = "un"
       Case "울"
            hantoeng = "ul"
       Case "움"
            hantoeng = "um"
       Case "웅"
            hantoeng = "ung"
       Case "워"
            hantoeng = "wo"
       Case "원"
            hantoeng = "won"
       Case "월"
            hantoeng = "wol"
       Case "위"
            hantoeng = "wi"
       Case "유"
            hantoeng = "yu"
       Case "육"
            hantoeng = "yuk"
       Case "윤"
            hantoeng = "yun"
       Case "율"
            hantoeng = "yul"
       Case "융"
            hantoeng = "yung"
       Case "윷"
            hantoeng = "yut"
       Case "으"
            hantoeng = "eu"
       Case "은"
            hantoeng = "eun"
       Case "을"
            hantoeng = "eul"
       Case "음"
            hantoeng = "eum"
       Case "읍"
            hantoeng = "eup"
       Case "응"
            hantoeng = "eung"
       Case "의"
            hantoeng = "ui"
       Case "이"
            hantoeng = "I"
       Case "익"
            hantoeng = "Ik"
       Case "인"
            hantoeng = "In"
       Case "일"
            hantoeng = "Il"
       Case "임"
            hantoeng = "Im"
       Case "입"
            hantoeng = "Ip"
       Case "잉"
            hantoeng = "Ing"
       Case "자"
            hantoeng = "ja"
       Case "작"
            hantoeng = "jak"
       Case "잔"
            hantoeng = "jan"
       Case "잠"
            hantoeng = "jam"
       Case "잡"
            hantoeng = "jap"
       Case "장"
            hantoeng = "jang"
       Case "재"
            hantoeng = "jae"
       Case "쟁"
            hantoeng = "jaeng"
       Case "저"
            hantoeng = "jeo"
       Case "적"
            hantoeng = "jeok"
       Case "전"
            hantoeng = "jeon"
       Case "절"
            hantoeng = "jeol"
       Case "점"
            hantoeng = "jeom"
       Case "접"
            hantoeng = "jeop"
       Case "정"
            hantoeng = "jeong"
       Case "제"
            hantoeng = "je"
       Case "조"
            hantoeng = "jo"
       Case "족"
            hantoeng = "jok"
       Case "존"
            hantoeng = "jon"
       Case "졸"
            hantoeng = "jol"
       Case "종"
            hantoeng = "jong"
       Case "좌"
            hantoeng = "jwa"
       Case "죄"
            hantoeng = "joe"
       Case "주"
            hantoeng = "ju"
       Case "죽"
            hantoeng = "juk"
       Case "준"
            hantoeng = "jun"
       Case "줄"
            hantoeng = "jul"
       Case "중"
            hantoeng = "jung"
       Case "쥐"
            hantoeng = "jwi"
       Case "즈"
            hantoeng = "jeu"
       Case "즉"
            hantoeng = "jeuk"
       Case "즐"
            hantoeng = "jeul"
       Case "즘"
            hantoeng = "jeum"
       Case "즙"
            hantoeng = "jeup"
       Case "증"
            hantoeng = "jeung"
       Case "지"
            hantoeng = "ji"
       Case "직"
            hantoeng = "jik"
       Case "진"
            hantoeng = "jin"
       Case "질"
            hantoeng = "jil"
       Case "짐"
            hantoeng = "jim"
       Case "집"
            hantoeng = "jip"
       Case "징"
            hantoeng = "jing"
       Case "짜"
            hantoeng = "jja"
       Case "째"
            hantoeng = "jjae"
       Case "쪼"
            hantoeng = "jjo"
       Case "찌"
            hantoeng = "jji"
       Case "차"
            hantoeng = "cha"
       Case "착"
            hantoeng = "chak"
       Case "찬"
            hantoeng = "chan"
       Case "찰"
            hantoeng = "chal"
       Case "참"
            hantoeng = "cham"
       Case "창"
            hantoeng = "chang"
       Case "채"
            hantoeng = "chae"
       Case "책"
            hantoeng = "chaek"
       Case "처"
            hantoeng = "cheo"
       Case "척"
            hantoeng = "cheok"
       Case "천"
            hantoeng = "cheon"
       Case "철"
            hantoeng = "cheol"
       Case "첨"
            hantoeng = "cheom"
       Case "첩"
            hantoeng = "cheop"
       Case "청"
            hantoeng = "cheong"
       Case "체"
            hantoeng = "che"
       Case "초"
            hantoeng = "cho"
       Case "촉"
            hantoeng = "chok"
       Case "촌"
            hantoeng = "chon"
       Case "총"
            hantoeng = "chong"
       Case "최"
            hantoeng = "choe"
       Case "추"
            hantoeng = "chu"
       Case "축"
            hantoeng = "chuk"
       Case "춘"
            hantoeng = "chun"
       Case "출"
            hantoeng = "chul"
       Case "춤"
            hantoeng = "chum"
       Case "충"
            hantoeng = "chung"
       Case "측"
            hantoeng = "cheuk"
       Case "측"
            hantoeng = "cheuk"
       Case "층"
            hantoeng = "cheung"
       Case "치"
            hantoeng = "chi"
       Case "칙"
            hantoeng = "chik"
       Case "친"
            hantoeng = "chin"
       Case "칠"
            hantoeng = "chil"
       Case "침"
            hantoeng = "chim"
       Case "칩"
            hantoeng = "chip"
       Case "칭"
            hantoeng = "ching"
       Case "코"
            hantoeng = "ko"
       Case "쾌"
            hantoeng = "kwae"
       Case "크"
            hantoeng = "keu"
       Case "큰"
            hantoeng = "keun"
       Case "키"
            hantoeng = "ki"
       Case "타"
            hantoeng = "ta"
       Case "탁"
            hantoeng = "tak"
       Case "탄"
            hantoeng = "tan"
       Case "탈"
            hantoeng = "tal"
       Case "탐"
            hantoeng = "tam"
       Case "탑"
            hantoeng = "tap"
       Case "탕"
            hantoeng = "tang"
       Case "태"
            hantoeng = "tae"
       Case "택"
            hantoeng = "taek"
       Case "탱"
            hantoeng = "taeng"
       Case "터"
            hantoeng = "teo"
       Case "테"
            hantoeng = "te"
       Case "토"
            hantoeng = "to"
       Case "톤"
            hantoeng = "ton"
       Case "톨"
            hantoeng = "tol"
       Case "통"
            hantoeng = "tong"
       Case "퇴"
            hantoeng = "toe"
       Case "투"
            hantoeng = "tu"
       Case "퉁"
            hantoeng = "tung"
       Case "튀"
            hantoeng = "twi"
       Case "트"
            hantoeng = "teu"
       Case "특"
            hantoeng = "teuk"
       Case "틈"
            hantoeng = "teum"
       Case "티"
            hantoeng = "ti"
       Case "파"
            hantoeng = "pa"
       Case "판"
            hantoeng = "pan"
       Case "팔"
            hantoeng = "pal"
       Case "패"
            hantoeng = "pae"
       Case "팽"
            hantoeng = "paeng"
       Case "퍼"
            hantoeng = "peo"
       Case "페"
            hantoeng = "pe"
       Case "펴"
            hantoeng = "pyeo"
       Case "편"
            hantoeng = "pyeon"
       Case "폄"
            hantoeng = "pyeom"
       Case "평"
            hantoeng = "pyeong"
       Case "폐"
            hantoeng = "pye"
       Case "포"
            hantoeng = "po"
       Case "폭"
            hantoeng = "pok"
       Case "표"
            hantoeng = "pyo"
       Case "푸"
            hantoeng = "pu"
       Case "품"
            hantoeng = "pum"
       Case "풍"
            hantoeng = "pung"
       Case "프"
            hantoeng = "peu"
       Case "피"
            hantoeng = "pi"
       Case "픽"
            hantoeng = "pik"
       Case "필"
            hantoeng = "pil"
       Case "핍"
            hantoeng = "pip"
       Case "하"
            hantoeng = "ha"
       Case "학"
            hantoeng = "hak"
       Case "한"
            hantoeng = "han"
       Case "할"
            hantoeng = "hal"
       Case "함"
            hantoeng = "ham"
       Case "합"
            hantoeng = "hap"
       Case "항"
            hantoeng = "hang"
       Case "해"
            hantoeng = "hae"
       Case "핵"
            hantoeng = "haek"
       Case "행"
            hantoeng = "haeng"
       Case "향"
            hantoeng = "hyang"
       Case "허"
            hantoeng = "heo"
       Case "헌"
            hantoeng = "heon"
       Case "험"
            hantoeng = "heom"
       Case "헤"
            hantoeng = "he"
       Case "혀"
            hantoeng = "hyeo"
       Case "혁"
            hantoeng = "hyeok"
       Case "현"
            hantoeng = "hyeon"
       Case "혈"
            hantoeng = "hyeol"
       Case "혐"
            hantoeng = "hyeom"
       Case "협"
            hantoeng = "hyeop"
       Case "형"
            hantoeng = "hyeong"
       Case "혜"
            hantoeng = "hye"
       Case "호"
            hantoeng = "ho"
       Case "혹"
            hantoeng = "hok"
       Case "혼"
            hantoeng = "hon"
       Case "홀"
            hantoeng = "hol"
       Case "홉"
            hantoeng = "hop"
       Case "홍"
            hantoeng = "hong"
       Case "화"
            hantoeng = "hwa"
       Case "확"
            hantoeng = "hwak"
       Case "환"
            hantoeng = "hwan"
       Case "활"
            hantoeng = "hwal"
       Case "황"
            hantoeng = "hwang"
       Case "홰"
            hantoeng = "hwae"
       Case "횃"
            hantoeng = "hwaet"
       Case "회"
            hantoeng = "hoe"
       Case "획"
            hantoeng = "hoek"
       Case "횡"
            hantoeng = "hoeng"
       Case "효"
            hantoeng = "hyo"
       Case "후"
            hantoeng = "hu"
       Case "훈"
            hantoeng = "hun"
       Case "훤"
            hantoeng = "hwon"
       Case "훼"
            hantoeng = "hwe"
       Case "휘"
            hantoeng = "hwi"
       Case "휴"
            hantoeng = "hyu"
       Case "휼"
            hantoeng = "hyul"
       Case "흉"
            hantoeng = "hyung"
       Case "흐"
            hantoeng = "heu"
       Case "흑"
            hantoeng = "heuk"
       Case "흔"
            hantoeng = "heun"
       Case "흘"
            hantoeng = "heul"
       Case "흠"
            hantoeng = "heum"
       Case "흡"
            hantoeng = "heup"
       Case "흥"
            hantoeng = "heung"
       Case "희"
            hantoeng = "hui"
       Case "흰"
            hantoeng = "huin"
       Case "히"
            hantoeng = "hi"
       Case "힘"
            hantoeng = "him"
       
       Case " "     '2010.01.27 이상은 추가
            hantoeng = " "
       Case Else
            hantoeng = "??"
            
       End Select

End Function


