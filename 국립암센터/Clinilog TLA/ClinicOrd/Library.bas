Attribute VB_Name = "Library"
Option Explicit

Declare Function ImmGetContext Lib "imm32.dll" (ByVal hwnd As Long) As Long
Declare Function ImmGetConversionStatus Lib "imm32.dll" (ByVal hIMC As Long, lpdw As Long, lpdw2 As Long) As Long
Declare Function ImmSetConversionStatus Lib "imm32.dll" (ByVal hIMC As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Type PatGen
    Birth As String
    Age As String
    Sex As String
End Type
Public gPatGen As PatGen

Public Function SetSpace(asStr As String, asLen As Integer, Optional asPos As Integer = 1) As String
'asPos = 1 : Left ����
'asPos = 2 : Right ���� ä���
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

Public Function SetChar(asStr As String, asLen As Integer, Optional asPos As Integer = 1, Optional asChr As String = " ") As String
'asPos = 1 : Left ����
'asPos = 2 : Right ���� ä���
    Dim sTmp As String
    Dim i As Integer
    
    sTmp = ""
    If Len(asStr) >= asLen Then
        SetChar = Left(asStr, asLen)
        Exit Function
    End If
    
    sTmp = asStr
    For i = 1 To asLen - Len(asStr)
        If asPos = 1 Then
            sTmp = asChr & sTmp
        Else
            sTmp = sTmp & asChr
        End If
    Next i
    
    SetChar = sTmp
End Function

Public Function ChangeDateFormat(ByVal asStr As String, Optional argV As String = "/") As String
    If Len(asStr) = 10 Then
        ChangeDateFormat = Left(asStr, 4) & argV & Mid(asStr, 6, 2) & argV & Mid(asStr, 9, 2)
    ElseIf Len(asStr) = 8 Then
        ChangeDateFormat = Left(asStr, 4) & argV & Mid(asStr, 5, 2) & argV & Mid(asStr, 7, 2)
    End If
End Function

Public Sub InsertRow(ByVal vasTable As Object, ByVal argRow As Integer)
'�������忡 Row �߰�
    vasTable.MaxRows = vasTable.MaxRows + 1
    vasTable.Row = argRow
    vasTable.Action = 7
End Sub

Public Sub DeleteRow(ByVal vasTable As Object, ByVal argRow1 As Integer, ByVal argRow2 As Integer)
'�������忡 Row ����
    vasTable.Row = argRow1
    vasTable.Row2 = argRow2
    vasTable.Col = 1
    vasTable.Col2 = vasTable.MaxCols
    vasTable.BlockMode = True
    vasTable.Action = 5
    vasTable.BlockMode = False
End Sub

Public Sub SelectFocus(ByRef argObj As Object)
'GetFocus �� Object���� Text�� ��ü ���� �ǰ� �Ѵ�.
    argObj.SelStart = 0
    argObj.SelLength = Len(argObj.Text)
End Sub

Public Sub SaveQuery(argSQL As String, Optional argFlag As Integer = 0)
'argSQL�� ������ ���Ϸ� ����
    Dim FilNum
    
    FilNum = FreeFile
    
    Open App.Path & "\QueryErr.txt" For Append As FilNum
    Print #FilNum, argSQL
    Close FilNum
End Sub

Public Sub SaveData(argSQL As String)
'argSQL�� ������ ���Ϸ� ����
    Dim FilNum
    
    FilNum = FreeFile
    
    If Dir(App.Path & "\Log", vbDirectory) = "" Then
        MkDir App.Path & "\Log"
    End If
    
    Open App.Path & "\Log\" & Format(Date, "yyyymmdd") & ".log" For Append As FilNum
    Print #FilNum, Time & " " & argSQL
    Close FilNum
End Sub

Public Function CR() As String
    CR = Chr(13) & Chr(10)
End Function

Public Function vasActiveCell(ByRef vasTable As Object, ByVal vasRow As Integer, ByVal vasCol As Integer) As Boolean
'Ư�� Cell ����
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    vasTable.Action = 0
End Function

Public Function GetCurRow(ByRef vasTable As Object) As Integer
'���� Active �� Row �����´�
    GetCurRow = vasTable.ActiveRow
End Function

Public Function GetCurCol(ByRef vasTable As Object) As Integer
'���� Active �� Col �����´�
    GetCurCol = vasTable.ActiveCol
End Function

Public Sub ClearSpread(ByRef vasTable As Object, Optional argStartRow As Long = 1, Optional argStartCol As Long = 0)
'vsSpread�� ������ Clear �Ѵ�.
    vasTable.Row = argStartRow
    vasTable.Col = argStartCol
    vasTable.Row2 = vasTable.DataRowCnt
    vasTable.Col2 = vasTable.DataColCnt
    vasTable.BlockMode = True
    vasTable.Action = 3
    vasTable.BlockMode = False
End Sub
Public Function SetText(ByRef vasTable As Object, ByVal SetStr As String, ByVal vasRow As Integer, ByVal vasCol As Integer) As Boolean
'vsSpread�� ����Ÿ �ֱ�
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    vasTable.Text = SetStr
End Function

Public Function GetText(ByRef vasTable As Object, ByVal vasRow As Integer, ByVal vasCol As Integer) As String
'vsSpread���� ����Ÿ ��������
    If vasRow < 0 Or vasCol < 0 Then
        Exit Function
    End If
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    GetText = vasTable.Text
End Function

Sub SetBackColor(asTable As vaSpread, ByVal asRow1 As Long, ByVal asRow2 As Long, ByVal asCol1 As Long, ByVal asCol2 As Long, asR As Variant, asG As Variant, asB As Variant)
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

Public Sub SetIME(h As Long, Toggle As Boolean)
'2003/07/06 �̻��� �߰�

'                 h:�� �ڵ�, Toggle:��/��(true/false)
'====================================================
'   �ѱ۷� ��ȯ    Call SetIME(Form1.hWnd, True)
'   ����� ��ȯ    Call SetIME(Form1.hWnd, False)
'====================================================
    Dim hIMC As Long
    Dim dwConversion As Long, dwSentence As Long
    Dim Temp As Long '
    
    hIMC = ImmGetContext(h)
    Temp = ImmGetConversionStatus(hIMC, dwConversion, dwSentence)
    If Toggle Then
        dwConversion = dwConversion Or 1
        Temp = ImmSetConversionStatus(hIMC, dwConversion, dwSentence)
    Else
        dwConversion = dwConversion And -2&
        Temp = ImmSetConversionStatus(hIMC, dwConversion, dwSentence)
    End If
End Sub

Public Function vasSort(ByRef vasTable As Object, ByVal key1 As Integer, Optional key2 As Integer = 0, Optional key3 As Integer = 0, Optional key4 As Integer = 0, Optional key5 As Integer = 0) As Boolean
'������ �κ��� ����
    vasTable.Row = 0
    vasTable.Col = 0
    vasTable.Row2 = vasTable.DataRowCnt
    vasTable.Col2 = vasTable.DataColCnt
'������ Row�� �ǽ�
    vasTable.SortBy = 2 'SS_SORT_BY_ROW
'���� Ű�� ����
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
'����
    vasTable.Action = 25 'SS_ACTION_SORT

    vasActiveCell vasTable, 1, 1
End Function

Public Sub DoSleep(Optional ByVal lMilliSec As Long = 0)
    'The DoSleep function allows other threads to have a time slice
    'and still keeps the main VB thread alive (since DPlay callbacks
    'run on separate threads outside of VB).
    Sleep lMilliSec
    DoEvents
End Sub

Public Sub CalSexAge(ByRef asPNRN As String, ByRef asCurDate As String)
    Dim sBirth As String
    Dim sStart As String
    Dim lAge As Long
    
    gPatGen.Sex = ""
    gPatGen.Age = ""
    gPatGen.Birth = ""
    
    If Mid(asPNRN, 1, 1) = "_" Or Mid(asPNRN, 1, 1) = "" Then
        Exit Sub
    End If
         
    asPNRN = Trim(asPNRN)
    
    If Len(asCurDate) = 8 Then
        asCurDate = Left(asCurDate, 4) & "-" & Mid(asCurDate, 5, 2) & "-" & Mid(asCurDate, 7, 2)
    End If
    
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
    
    sBirth = sBirth & Mid(asPNRN, 1, 2) '& "/" & Mid(asPNRN, 3, 2) & "/" & Mid(asPNRN, 5, 2)
    If Mid(asPNRN, 3, 2) = "00" Then
        sBirth = sBirth & "/01"
    Else
        sBirth = sBirth & "/" & Mid(asPNRN, 3, 2)
    End If
    If Mid(asPNRN, 5, 2) = "00" Then
        sBirth = sBirth & "/01"
    Else
        sBirth = sBirth & "/" & Mid(asPNRN, 5, 2)
    End If
    
    If IsDate(sBirth) Then
        gPatGen.Birth = sBirth
        
        lAge = DateDiff("yyyy", sBirth, asCurDate)
        If lAge < 1 Then
            If DateDiff("d", sBirth, asCurDate) <= 15 Then
                gPatGen.Age = DateDiff("d", sBirth, asCurDate) & "��"
            Else
                gPatGen.Age = lAge
            End If
        Else
            gPatGen.Age = lAge
        End If
    Else
        gPatGen.Birth = ""
        gPatGen.Age = "0"
    End If
    
End Sub

Public Function uLen(Str As String) As Long
' ���ڿ����� ���ϱ� �ѱ�(2Byte)  (Str:���ڿ�)
    Dim i As Long
    Dim chLen As Long
    For i = 1 To Len(Str)
        If Asc(Mid$(Str, i, 1)) < 0 Then
            chLen = chLen + 2
        Else
            chLen = chLen + 1
        End If
    Next i
    uLen = chLen
End Function

Public Function uMid(Str As String, Stp As Long, Nct As Long) As String
' ���ڿ� ���� �ѱ�(2Byte)  (Str:���ڿ�, Stp:������ġ, Nct:���ⰳ��)
    Dim i As Long
    Dim chLen As Long
    Dim RetStr As String
    For i = 1 To Len(Str)
        If Asc(Mid$(Str, i, 1)) < 0 Then
            chLen = chLen + 2
        Else
            chLen = chLen + 1
        End If
        If (chLen >= Stp) And (chLen <= Stp + Nct - 1) Then
            RetStr = RetStr & Mid(Str, i, 1)
        End If
    Next i
     uMid = RetStr
End Function

Public Function LPad(pText As String, pLength As Long, pChar As String) As String
'���� ������ Ư�����ڷ� ä��(ORACLE LPAD�� ����)
'pText : �⺻ ���ڿ�
'pLength : ��ü ����
'RETURN CODE : ���� ������ Ư�����ڷ� ä�� ���ڿ�
    If pLength = -1 Then
        LPad = pText
    Else
        If uLen(pText) <= pLength Then
            LPad = String(pLength - uLen(pText), pChar) & pText
        Else
            LPad = IIf(uLen(uMid(pText, 1, pLength)) = pLength - 1, Space(1), "") & uMid(pText, 1, pLength)
            'MsgBox "Too Large Text for LPad Function", vbCritical
        End If
    End If
End Function

Public Function Conv_Kor_Eng(ByVal asName As String) As String
    Dim sName As String
    Dim i As Integer
    
    sName = ""
    
    For i = 1 To Len(asName)
        sName = sName & hantoeng(Mid(asName, i, 1))
    
        If i = 1 Then
            If sName = "I" Then
                sName = "lee"
            End If
        End If
    
    Next i
    
    Conv_Kor_Eng = Trim(sName)
End Function

Public Function hantoeng(onehan As String) As String
       Select Case onehan
       Case "��"
            hantoeng = "ga"
       Case "��"
            hantoeng = "gak"
       Case "��"
            hantoeng = "gan"
       Case "��"
            hantoeng = "gal"
       Case "��"
            hantoeng = "gam"
       Case "��"
            hantoeng = "gap"
       Case "��"
            hantoeng = "gat"
       Case "��"
            hantoeng = "gang"
       Case "��"
            hantoeng = "gae"
       Case "��"
            hantoeng = "gaek"
       Case "��"
            hantoeng = "geo"
       Case "��"
            hantoeng = "geon"
       Case "��"
            hantoeng = "geol"
       Case "��"
            hantoeng = "geom"
       Case "��"
            hantoeng = "geop"
       Case "��"
            hantoeng = "ge"
       Case "��"
            hantoeng = "gyeo"
       Case "��"
            hantoeng = "gyeok"
       Case "��"
            hantoeng = "gyeon"
       Case "��"
            hantoeng = "gyeol"
       Case "��"
            hantoeng = "gyeom"
       Case "��"
            hantoeng = "gyeop"
       Case "��"
            hantoeng = "gyeong"
       Case "��"
            hantoeng = "gye"
       Case "��"
            hantoeng = "go"
       Case "��"
            hantoeng = "gok"
       Case "��"
            hantoeng = "gon"
       Case "��"
            hantoeng = "gol"
       Case "��"
            hantoeng = "got"
       Case "��"
            hantoeng = "gong"
       Case "��"
            hantoeng = "got"
       Case "��"
            hantoeng = "gwa"
       Case "��"
            hantoeng = "gwak"
       Case "��"
            hantoeng = "gwan"
       Case "��"
            hantoeng = "gwal"
       Case "��"
            hantoeng = "gwang"
       Case "��"
            hantoeng = "gwae"
       Case "��"
            hantoeng = "goe"
       Case "��"
            hantoeng = "goeng"
       Case "��"
            hantoeng = "gyo"
       Case "��"
            hantoeng = "gu"
       Case "��"
            hantoeng = "guk"
       Case "��"
            hantoeng = "gun"
       Case "��"
            hantoeng = "gul"
       Case "��"
            hantoeng = "gut"
       Case "��"
            hantoeng = "gung"
       Case "��"
            hantoeng = "gwon"
       Case "��"
            hantoeng = "gwol"
       Case "��"
            hantoeng = "gwi"
       Case "��"
            hantoeng = "gyu"
       Case "��"
            hantoeng = "gyun"
       Case "��"
            hantoeng = "gyul"
       Case "��"
            hantoeng = "geu"
       Case "��"
            hantoeng = "geuk"
       Case "��"
            hantoeng = "geun"
       Case "��"
            hantoeng = "geul"
       Case "��"
            hantoeng = "geul"
       Case "��"
            hantoeng = "geum"
       Case "��"
            hantoeng = "geup"
       Case "��"
            hantoeng = "geung"
       Case "��"
            hantoeng = "gi"
       Case "��"
            hantoeng = "gin"
       Case "��"
            hantoeng = "gil"
       Case "��"
            hantoeng = "gim"
       Case "��"
            hantoeng = "kka"
       Case "��"
            hantoeng = "kkae"
       Case "��"
            hantoeng = "kko"
       Case "��"
            hantoeng = "kkok"
       Case "��"
            hantoeng = "kkot"
       Case "��"
            hantoeng = "kkoe"
       Case "��"
            hantoeng = "kku"
       Case "��"
            hantoeng = "kkum"
       Case "��"
            hantoeng = "kkeut"
       Case "��"
            hantoeng = "kki"
       Case "��"
            hantoeng = "na"
       Case "��"
            hantoeng = "nak"
       Case "��"
            hantoeng = "nan"
       Case "��"
            hantoeng = "nal"
       Case "��"
            hantoeng = "nam"
       Case "��"
            hantoeng = "nap"
       Case "��"
            hantoeng = "nang"
       Case "��"
            hantoeng = "nae"
       Case "��"
            hantoeng = "naeng"
       Case "��"
            hantoeng = "neo"
       Case "��"
            hantoeng = "neol"
       Case "��"
            hantoeng = "ne"
       Case "��"
            hantoeng = "nyeo"
       Case "��"
            hantoeng = "nyeok"
       Case "��"
            hantoeng = "nyeon"
       Case "��"
            hantoeng = "nyeom"
       Case "��"
            hantoeng = "nyeong"
       Case "��"
            hantoeng = "no"
       Case "��"
            hantoeng = "nok"
       Case "��"
            hantoeng = "non"
       Case "��"
            hantoeng = "nol"
       Case "��"
            hantoeng = "nong"
       Case "��"
            hantoeng = "noe"
       Case "��"
            hantoeng = "nu"
       Case "��"
            hantoeng = "nun"
       Case "��"
            hantoeng = "nul"
       Case "��"
            hantoeng = "neu"
       Case "��"
            hantoeng = "neuk"
       Case "��"
            hantoeng = "neum"
       Case "��"
            hantoeng = "neung"
       Case "��"
            hantoeng = "nui"
       Case "��"
            hantoeng = "ni"
       Case "��"
            hantoeng = "nik"
       Case "��"
            hantoeng = "nin"
       Case "��"
            hantoeng = "nil"
       Case "��"
            hantoeng = "nim"
       Case "��"
            hantoeng = "da"
       Case "��"
            hantoeng = "dan"
       Case "��"
            hantoeng = "dal"
       Case "��"
            hantoeng = "dam"
       Case "��"
            hantoeng = "dap"
       Case "��"
            hantoeng = "dang"
       Case "��"
            hantoeng = "dae"
       Case "��"
            hantoeng = "daek"
       Case "��"
            hantoeng = "deo"
       Case "��"
            hantoeng = "deok"
       Case "��"
            hantoeng = "do"
       Case "��"
            hantoeng = "dok"
       Case "��"
            hantoeng = "don"
       Case "��"
            hantoeng = "dol"
       Case "��"
            hantoeng = "dong"
       Case "��"
            hantoeng = "dwae"
       Case "��"
            hantoeng = "doe"
       Case "��"
            hantoeng = "doen"
       Case "��"
            hantoeng = "du"
       Case "��"
            hantoeng = "duk"
       Case "��"
            hantoeng = "dun"
       Case "��"
            hantoeng = "dwi"
       Case "��"
            hantoeng = "deu"
       Case "��"
            hantoeng = "deuk"
       Case "��"
            hantoeng = "deul"
       Case "��"
            hantoeng = "deung"
       Case "��"
            hantoeng = "di"
       Case "��"
            hantoeng = "tta"
       Case "��"
            hantoeng = "ttang"
       Case "��"
            hantoeng = "ttae"
       Case "��"
            hantoeng = "tto"
       Case "��"
            hantoeng = "ttu"
       Case "��"
            hantoeng = "ttuk"
       Case "��"
            hantoeng = "tteu"
       Case "��"
            hantoeng = "tti"
       Case "��"
            hantoeng = "ra"
       Case "��"
            hantoeng = "rak"
       Case "��"
            hantoeng = "ran"
       Case "��"
            hantoeng = "ram"
       Case "��"
            hantoeng = "rang"
       Case "��"
            hantoeng = "rae"
       Case "��"
            hantoeng = "raeng"
       Case "��"
            hantoeng = "ryang"
       Case "��"
            hantoeng = "reong"
       Case "��"
            hantoeng = "re"
       Case "��"
            hantoeng = "ryeo"
       Case "��"
            hantoeng = "ryeok"
       Case "��"
            hantoeng = "ryeon"
       Case "��"
            hantoeng = "ryeol"
       Case "��"
            hantoeng = "ryeom"
       Case "��"
            hantoeng = "ryeop"
       Case "��"
            hantoeng = "ryeong"
       Case "��"
            hantoeng = "rye"
       Case "��"
            hantoeng = "ro"
       Case "��"
            hantoeng = "rok"
       Case "��"
            hantoeng = "ron"
       Case "��"
            hantoeng = "rong"
       Case "��"
            hantoeng = "roe"
       Case "��"
            hantoeng = "ryo"
       Case "��"
            hantoeng = "ryong"
       Case "��"
            hantoeng = "ru"
       Case "��"
            hantoeng = "ryu"
       Case "��"
            hantoeng = "ryuk"
       Case "��"
            hantoeng = "ryun"
       Case "��"
            hantoeng = "ryul"
       Case "��"
            hantoeng = "ryung"
       Case "��"
            hantoeng = "reu"
       Case "��"
            hantoeng = "reuk"
       Case "��"
            hantoeng = "reun"
       Case "��"
            hantoeng = "reum"
       Case "��"
            hantoeng = "reung"
       Case "��"
            hantoeng = "ri"
       Case "��"
            hantoeng = "rin"
       Case "��"
            hantoeng = "rim"
       Case "��"
            hantoeng = "rip"
       Case "��"
            hantoeng = "ma"
       Case "��"
            hantoeng = "mak"
       Case "��"
            hantoeng = "man"
       Case "��"
            hantoeng = "mal"
       Case "��"
            hantoeng = "mang"
       Case "��"
            hantoeng = "mae"
       Case "��"
            hantoeng = "maek"
       Case "��"
            hantoeng = "maen"
       Case "��"
            hantoeng = "maeng"
       Case "��"
            hantoeng = "meo"
       Case "��"
            hantoeng = "meok"
       Case "��"
            hantoeng = "me"
       Case "��"
            hantoeng = "myeo"
       Case "��"
            hantoeng = "myeok"
       Case "��"
            hantoeng = "myeon"
       Case "��"
            hantoeng = "myeol"
       Case "��"
            hantoeng = "myeong"
       Case "��"
            hantoeng = "mo"
       Case "��"
            hantoeng = "mok"
       Case "��"
            hantoeng = "mol"
       Case "��"
            hantoeng = "mot"
       Case "��"
            hantoeng = "mong"
       Case "��"
            hantoeng = "moe"
       Case "��"
            hantoeng = "myo"
       Case "��"
            hantoeng = "mu"
       Case "��"
            hantoeng = "muk"
       Case "��"
            hantoeng = "mun"
       Case "��"
            hantoeng = "mul"
       Case "��"
            hantoeng = "meu"
       Case "��"
            hantoeng = "mi"
       Case "��"
            hantoeng = "min"
       Case "��"
            hantoeng = "mil"
       Case "��"
            hantoeng = "ba"
       Case "��"
            hantoeng = "bak"
       Case "��"
            hantoeng = "ban"
       Case "��"
            hantoeng = "bal"
       Case "��"
            hantoeng = "bap"
       Case "��"
            hantoeng = "bang"
       Case "��"
            hantoeng = "bae"
       Case "��"
            hantoeng = "baek"
       Case "��"
            hantoeng = "baem"
       Case "��"
            hantoeng = "beo"
       Case "��"
            hantoeng = "beon"
       Case "��"
            hantoeng = "beol"
       Case "��"
            hantoeng = "beom"
       Case "��"
            hantoeng = "beop"
       Case "��"
            hantoeng = "byeo"
       Case "��"
            hantoeng = "byeok"
       Case "��"
            hantoeng = "byeon"
       Case "��"
            hantoeng = "byeol"
       Case "��"
            hantoeng = "byeong"
       Case "��"
            hantoeng = "bo"
       Case "��"
            hantoeng = "bok"
       Case "��"
            hantoeng = "bon"
       Case "��"
            hantoeng = "bong"
       Case "��"
            hantoeng = "bu"
       Case "��"
            hantoeng = "buk"
       Case "��"
            hantoeng = "bun"
       Case "��"
            hantoeng = "bul"
       Case "��"
            hantoeng = "bung"
       Case "��"
            hantoeng = "bi"
       Case "��"
            hantoeng = "bin"
       Case "��"
            hantoeng = "bil"
       Case "��"
            hantoeng = "bim"
       Case "��"
            hantoeng = "bing"
       Case "��"
            hantoeng = "ppa"
       Case "��"
            hantoeng = "ppae"
       Case "��"
            hantoeng = "ppeo"
       Case "��"
            hantoeng = "ppo"
       Case "��"
            hantoeng = "ppu"
       Case "��"
            hantoeng = "ppeu"
       Case "��"
            hantoeng = "ppi"
       Case "��"
            hantoeng = "sa"
       Case "��"
            hantoeng = "sak"
       Case "��"
            hantoeng = "san"
       Case "��"
            hantoeng = "sal"
       Case "��"
            hantoeng = "sam"
       Case "��"
            hantoeng = "sap"
       Case "��"
            hantoeng = "sang"
       Case "��"
            hantoeng = "sat"
       Case "��"
            hantoeng = "sae"
       Case "��"
            hantoeng = "saek"
       Case "��"
            hantoeng = "saeng"
       Case "��"
            hantoeng = "seo"
       Case "��"
            hantoeng = "seok"
       Case "��"
            hantoeng = "seon"
       Case "��"
            hantoeng = "seol"
       Case "��"
            hantoeng = "seom"
       Case "��"
            hantoeng = "seop"
       Case "��"
            hantoeng = "seong"
       Case "��"
            hantoeng = "se"
       Case "��"
            hantoeng = "syeo"
       Case "��"
            hantoeng = "so"
       Case "��"
            hantoeng = "syo"
       Case "��"
            hantoeng = "sok"
       Case "��"
            hantoeng = "son"
       Case "��"
            hantoeng = "sol"
       Case "��"
            hantoeng = "som"
       Case "��"
            hantoeng = "sot"
       Case "��"
            hantoeng = "song"
       Case "��"
            hantoeng = "swae"
       Case "��"
            hantoeng = "soe"
       Case "��"
            hantoeng = "su"
       Case "��"
            hantoeng = "suk"
       Case "��"
            hantoeng = "sun"
       Case "��"
            hantoeng = "sul"
       Case "��"
            hantoeng = "sum"
       Case "��"
            hantoeng = "sung"
       Case "��"
            hantoeng = "swi"
       Case "��"
            hantoeng = "seu"
       Case "��"
            hantoeng = "seul"
       Case "��"
            hantoeng = "seum"
       Case "��"
            hantoeng = "seup"
       Case "��"
            hantoeng = "seung"
       Case "��"
            hantoeng = "si"
       Case "��"
            hantoeng = "sik"
       Case "��"
            hantoeng = "sin"
       Case "��"
            hantoeng = "sil"
       Case "��"
            hantoeng = "sim"
       Case "��"
            hantoeng = "sip"
       Case "��"
            hantoeng = "sing"
       Case "��"
            hantoeng = "ssa"
       Case "��"
            hantoeng = "ssang"
       Case "��"
            hantoeng = "ssae"
       Case "��"
            hantoeng = "sso"
       Case "��"
            hantoeng = "ssuk"
       Case "��"
            hantoeng = "ssi"
       Case "��"
            hantoeng = "a"
       Case "��"
            hantoeng = "ak"
       Case "��"
            hantoeng = "an"
       Case "��"
            hantoeng = "al"
       Case "��"
            hantoeng = "am"
       Case "��"
            hantoeng = "ap"
       Case "��"
            hantoeng = "ang"
       Case "��"
            hantoeng = "ap"
       Case "��"
            hantoeng = "ae"
       Case "��"
            hantoeng = "aek"
       Case "��"
            hantoeng = "aeng"
       Case "��"
            hantoeng = "ya"
       Case "��"
            hantoeng = "yak"
       Case "��"
            hantoeng = "yan"
       Case "��"
            hantoeng = "yang"
       Case "��"
            hantoeng = "eo"
       Case "��"
            hantoeng = "eok"
       Case "��"
            hantoeng = "eon"
       Case "��"
            hantoeng = "eol"
       Case "��"
            hantoeng = "eom"
       Case "��"
            hantoeng = "eop"
       Case "��"
            hantoeng = "e"
       Case "��"
            hantoeng = "el"
       Case "��"
            hantoeng = "yeo"
       Case "��"
            hantoeng = "yeok"
       Case "��"
            hantoeng = "yeon"
       Case "��"
            hantoeng = "yeol"
       Case "��"
            hantoeng = "yeom"
       Case "��"
            hantoeng = "yeop"
       Case "��"
            hantoeng = "yeong"
       Case "��"
            hantoeng = "ye"
       Case "��"
            hantoeng = "o"
       Case "��"
            hantoeng = "ok"
       Case "��"
            hantoeng = "on"
       Case "��"
            hantoeng = "ol"
       Case "��"
            hantoeng = "om"
       Case "��"
            hantoeng = "ong"
       Case "��"
            hantoeng = "wa"
       Case "��"
            hantoeng = "wan"
       Case "��"
            hantoeng = "wal"
       Case "��"
            hantoeng = "wang"
       Case "��"
            hantoeng = "wae"
       Case "��"
            hantoeng = "oe"
       Case "��"
            hantoeng = "oen"
       Case "��"
            hantoeng = "yo"
       Case "��"
            hantoeng = "yok"
       Case "��"
            hantoeng = "yong"
       Case "��"
            hantoeng = "u"
       Case "��"
            hantoeng = "uk"
       Case "��"
            hantoeng = "un"
       Case "��"
            hantoeng = "ul"
       Case "��"
            hantoeng = "um"
       Case "��"
            hantoeng = "ung"
       Case "��"
            hantoeng = "wo"
       Case "��"
            hantoeng = "won"
       Case "��"
            hantoeng = "wol"
       Case "��"
            hantoeng = "wi"
       Case "��"
            hantoeng = "yu"
       Case "��"
            hantoeng = "yuk"
       Case "��"
            hantoeng = "yun"
       Case "��"
            hantoeng = "yul"
       Case "��"
            hantoeng = "yung"
       Case "��"
            hantoeng = "yut"
       Case "��"
            hantoeng = "eu"
       Case "��"
            hantoeng = "eun"
       Case "��"
            hantoeng = "eul"
       Case "��"
            hantoeng = "eum"
       Case "��"
            hantoeng = "eup"
       Case "��"
            hantoeng = "eung"
       Case "��"
            hantoeng = "ui"
       Case "��"
            hantoeng = "I"
       Case "��"
            hantoeng = "Ik"
       Case "��"
            hantoeng = "In"
       Case "��"
            hantoeng = "Il"
       Case "��"
            hantoeng = "Im"
       Case "��"
            hantoeng = "Ip"
       Case "��"
            hantoeng = "Ing"
       Case "��"
            hantoeng = "ja"
       Case "��"
            hantoeng = "jak"
       Case "��"
            hantoeng = "jan"
       Case "��"
            hantoeng = "jam"
       Case "��"
            hantoeng = "jap"
       Case "��"
            hantoeng = "jang"
       Case "��"
            hantoeng = "jae"
       Case "��"
            hantoeng = "jaeng"
       Case "��"
            hantoeng = "jeo"
       Case "��"
            hantoeng = "jeok"
       Case "��"
            hantoeng = "jeon"
       Case "��"
            hantoeng = "jeol"
       Case "��"
            hantoeng = "jeom"
       Case "��"
            hantoeng = "jeop"
       Case "��"
            hantoeng = "jeong"
       Case "��"
            hantoeng = "je"
       Case "��"
            hantoeng = "jo"
       Case "��"
            hantoeng = "jok"
       Case "��"
            hantoeng = "jon"
       Case "��"
            hantoeng = "jol"
       Case "��"
            hantoeng = "jong"
       Case "��"
            hantoeng = "jwa"
       Case "��"
            hantoeng = "joe"
       Case "��"
            hantoeng = "ju"
       Case "��"
            hantoeng = "juk"
       Case "��"
            hantoeng = "jun"
       Case "��"
            hantoeng = "jul"
       Case "��"
            hantoeng = "jung"
       Case "��"
            hantoeng = "jwi"
       Case "��"
            hantoeng = "jeu"
       Case "��"
            hantoeng = "jeuk"
       Case "��"
            hantoeng = "jeul"
       Case "��"
            hantoeng = "jeum"
       Case "��"
            hantoeng = "jeup"
       Case "��"
            hantoeng = "jeung"
       Case "��"
            hantoeng = "ji"
       Case "��"
            hantoeng = "jik"
       Case "��"
            hantoeng = "jin"
       Case "��"
            hantoeng = "jil"
       Case "��"
            hantoeng = "jim"
       Case "��"
            hantoeng = "jip"
       Case "¡"
            hantoeng = "jing"
       Case "¥"
            hantoeng = "jja"
       Case "°"
            hantoeng = "jjae"
       Case "��"
            hantoeng = "jjo"
       Case "��"
            hantoeng = "jji"
       Case "��"
            hantoeng = "cha"
       Case "��"
            hantoeng = "chak"
       Case "��"
            hantoeng = "chan"
       Case "��"
            hantoeng = "chal"
       Case "��"
            hantoeng = "cham"
       Case "â"
            hantoeng = "chang"
       Case "ä"
            hantoeng = "chae"
       Case "å"
            hantoeng = "chaek"
       Case "ó"
            hantoeng = "cheo"
       Case "ô"
            hantoeng = "cheok"
       Case "õ"
            hantoeng = "cheon"
       Case "ö"
            hantoeng = "cheol"
       Case "÷"
            hantoeng = "cheom"
       Case "ø"
            hantoeng = "cheop"
       Case "û"
            hantoeng = "cheong"
       Case "ü"
            hantoeng = "che"
       Case "��"
            hantoeng = "cho"
       Case "��"
            hantoeng = "chok"
       Case "��"
            hantoeng = "chon"
       Case "��"
            hantoeng = "chong"
       Case "��"
            hantoeng = "choe"
       Case "��"
            hantoeng = "chu"
       Case "��"
            hantoeng = "chuk"
       Case "��"
            hantoeng = "chun"
       Case "��"
            hantoeng = "chul"
       Case "��"
            hantoeng = "chum"
       Case "��"
            hantoeng = "chung"
       Case "��"
            hantoeng = "cheuk"
       Case "��"
            hantoeng = "cheuk"
       Case "��"
            hantoeng = "cheung"
       Case "ġ"
            hantoeng = "chi"
       Case "Ģ"
            hantoeng = "chik"
       Case "ģ"
            hantoeng = "chin"
       Case "ĥ"
            hantoeng = "chil"
       Case "ħ"
            hantoeng = "chim"
       Case "Ĩ"
            hantoeng = "chip"
       Case "Ī"
            hantoeng = "ching"
       Case "��"
            hantoeng = "ko"
       Case "��"
            hantoeng = "kwae"
       Case "ũ"
            hantoeng = "keu"
       Case "ū"
            hantoeng = "keun"
       Case "Ű"
            hantoeng = "ki"
       Case "Ÿ"
            hantoeng = "ta"
       Case "Ź"
            hantoeng = "tak"
       Case "ź"
            hantoeng = "tan"
       Case "Ż"
            hantoeng = "tal"
       Case "Ž"
            hantoeng = "tam"
       Case "ž"
            hantoeng = "tap"
       Case "��"
            hantoeng = "tang"
       Case "��"
            hantoeng = "tae"
       Case "��"
            hantoeng = "taek"
       Case "��"
            hantoeng = "taeng"
       Case "��"
            hantoeng = "teo"
       Case "��"
            hantoeng = "te"
       Case "��"
            hantoeng = "to"
       Case "��"
            hantoeng = "ton"
       Case "��"
            hantoeng = "tol"
       Case "��"
            hantoeng = "tong"
       Case "��"
            hantoeng = "toe"
       Case "��"
            hantoeng = "tu"
       Case "��"
            hantoeng = "tung"
       Case "Ƣ"
            hantoeng = "twi"
       Case "Ʈ"
            hantoeng = "teu"
       Case "Ư"
            hantoeng = "teuk"
       Case "ƴ"
            hantoeng = "teum"
       Case "Ƽ"
            hantoeng = "ti"
       Case "��"
            hantoeng = "pa"
       Case "��"
            hantoeng = "pan"
       Case "��"
            hantoeng = "pal"
       Case "��"
            hantoeng = "pae"
       Case "��"
            hantoeng = "paeng"
       Case "��"
            hantoeng = "peo"
       Case "��"
            hantoeng = "pe"
       Case "��"
            hantoeng = "pyeo"
       Case "��"
            hantoeng = "pyeon"
       Case "��"
            hantoeng = "pyeom"
       Case "��"
            hantoeng = "pyeong"
       Case "��"
            hantoeng = "pye"
       Case "��"
            hantoeng = "po"
       Case "��"
            hantoeng = "pok"
       Case "ǥ"
            hantoeng = "pyo"
       Case "Ǫ"
            hantoeng = "pu"
       Case "ǰ"
            hantoeng = "pum"
       Case "ǳ"
            hantoeng = "pung"
       Case "��"
            hantoeng = "peu"
       Case "��"
            hantoeng = "pi"
       Case "��"
            hantoeng = "pik"
       Case "��"
            hantoeng = "pil"
       Case "��"
            hantoeng = "pip"
       Case "��"
            hantoeng = "ha"
       Case "��"
            hantoeng = "hak"
       Case "��"
            hantoeng = "han"
       Case "��"
            hantoeng = "hal"
       Case "��"
            hantoeng = "ham"
       Case "��"
            hantoeng = "hap"
       Case "��"
            hantoeng = "hang"
       Case "��"
            hantoeng = "hae"
       Case "��"
            hantoeng = "haek"
       Case "��"
            hantoeng = "haeng"
       Case "��"
            hantoeng = "hyang"
       Case "��"
            hantoeng = "heo"
       Case "��"
            hantoeng = "heon"
       Case "��"
            hantoeng = "heom"
       Case "��"
            hantoeng = "he"
       Case "��"
            hantoeng = "hyeo"
       Case "��"
            hantoeng = "hyeok"
       Case "��"
            hantoeng = "hyeon"
       Case "��"
            hantoeng = "hyeol"
       Case "��"
            hantoeng = "hyeom"
       Case "��"
            hantoeng = "hyeop"
       Case "��"
            hantoeng = "hyeong"
       Case "��"
            hantoeng = "hye"
       Case "ȣ"
            hantoeng = "ho"
       Case "Ȥ"
            hantoeng = "hok"
       Case "ȥ"
            hantoeng = "hon"
       Case "Ȧ"
            hantoeng = "hol"
       Case "ȩ"
            hantoeng = "hop"
       Case "ȫ"
            hantoeng = "hong"
       Case "ȭ"
            hantoeng = "hwa"
       Case "Ȯ"
            hantoeng = "hwak"
       Case "ȯ"
            hantoeng = "hwan"
       Case "Ȱ"
            hantoeng = "hwal"
       Case "Ȳ"
            hantoeng = "hwang"
       Case "ȳ"
            hantoeng = "hwae"
       Case "ȶ"
            hantoeng = "hwaet"
       Case "ȸ"
            hantoeng = "hoe"
       Case "ȹ"
            hantoeng = "hoek"
       Case "Ⱦ"
            hantoeng = "hoeng"
       Case "ȿ"
            hantoeng = "hyo"
       Case "��"
            hantoeng = "hu"
       Case "��"
            hantoeng = "hun"
       Case "��"
            hantoeng = "hwon"
       Case "��"
            hantoeng = "hwe"
       Case "��"
            hantoeng = "hwi"
       Case "��"
            hantoeng = "hyu"
       Case "��"
            hantoeng = "hyul"
       Case "��"
            hantoeng = "hyung"
       Case "��"
            hantoeng = "heu"
       Case "��"
            hantoeng = "heuk"
       Case "��"
            hantoeng = "heun"
       Case "��"
            hantoeng = "heul"
       Case "��"
            hantoeng = "heum"
       Case "��"
            hantoeng = "heup"
       Case "��"
            hantoeng = "heung"
       Case "��"
            hantoeng = "hui"
       Case "��"
            hantoeng = "huin"
       Case "��"
            hantoeng = "hi"
       Case "��"
            hantoeng = "him"
       
       Case Else
            hantoeng = "??"
            
       End Select

End Function

