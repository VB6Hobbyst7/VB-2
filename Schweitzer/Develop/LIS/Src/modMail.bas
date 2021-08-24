Attribute VB_Name = "modMailing"
'Option Explicit
'
'Public Const conMailLongDate = 0
'Public Const conMailListView = 1
'
'Public Const vbRecipTypeTo = 1
'Public Const vbRecipTypeCc = 2
'
'Public Const conSSCount = 5
'Public Const conUnreadMessage = "*"     ' 읽지 않은 메시지를 나타내는 문자열에 대한 상수
'
'Public Const vbMessageFetch = 1
'Public Const vbMessageSendDlg = 2
'Public Const vbMessageSend = 3
'Public Const vbMessageSaveMsg = 4
'Public Const vbMessageCopy = 5
'Public Const vbMessageCompose = 6
'Public Const vbMessageReply = 7
'Public Const vbMessageReplyAll = 8
'Public Const vbMessageForward = 9
'Public Const vbMessageDelete = 10
'Public Const vbMessageShowAdBook = 11
'Public Const vbMessageShowDetails = 12
'Public Const vbMessageResolveName = 13
'Public Const vbRecipientDelete = 14
'Public Const vbAttachmentDelete = 15
'
'Public Const vbAttachTypeData = 0
'Public Const vbAttachTypeEOLE = 1
'Public Const vbAttachTypeSOLE = 2
'
'Public Const conHigh = 2100
'Public Const conLow = 1600
'
'Public currentSSRow As Integer
'
'Public currentRCIndex As Integer
'Public UnRead As Integer
'Public SendWithMapi As Integer
'Public ReturnRequest As Integer
'Public OptionType As Integer
'
'Function DateFROMMapiDate$(ByVal S$, wFormat%)
'' 이 프로시저는 MAPI 날짜의 서식 메시지를 보기 위한 두 서식 중 하나로 정합니다.
'
'    Y$ = Mid(S$, 1, 4)
'    M$ = Mid$(S$, 6, 2)
'    D$ = Mid$(S$, 9, 2)
'    T$ = Mid$(S$, 12)
'    Ds# = DateValue(M$ + "/" + D$ + "/" + Y$) + TimeValue(T$)
'
'    Select Case wFormat
'        Case conMailLongDate
'            F$ = "dddd, mmmm d, yyyy, h:mmAM/PM"
'        Case conMailListView
'            F$ = "mm/dd/yy hh:mm"
'    End Select
'
'    DateFROMMapiDate = Format$(Ds#, F$)
'
'End Function
'
'Function GetHeader(MSG As Control) As String
'Dim Header As String
'Dim CR As String
'
'    CR = Chr$(13) + Chr$(10)
'
'    Header = String$(25, "-") + CR
'    Header = Header + "폼: " + MSG.MsgOrigDisplayName + CR
'    Header = Header + "받는이: " + GetRCList(MSG, vbRecipTypeTo) + CR
'    Header = Header + "참조: " + GetRCList(MSG, vbRecipTypeCc) + CR
'    Header = Header + "제목: " + MSG.MsgSubject + CR
'    Header = Header + "보낸 날짜: " + DateFROMMapiDate$(MSG.MsgDateReceived, conMailLongDate) + CR + CR
'
'    GetHeader = Header
'
'End Function
'
'Function GetRCList(MSG As Control, RCType As Integer) As String
''받는 사람의 목록이 있는 경우 이 함수는 다음 유형에서 지정한
''형식으로 받는 사람의 목록을 반환합니다. (사람 1;사람 2;사람 3)
'
'    For i = 0 To MSG.RecipCount - 1
'        MSG.RecipIndex = i
'        If RCType = MSG.RecipType Then
'            A$ = A$ + ";" + MSG.RecipDisplayName
'        End If
'    Next i
'
'    If A$ <> "" Then
'        A$ = Mid$(A$, 2)  '앞쪽 ";"을 지웁니다.
'    End If
'
'    GetRCList = A$
'
'End Function
'
'Sub UpdateRecips()
'' 이 프로시저는 올바른 편집 필드와 받는 사람 정보를 새로 고칩니다.
'    txtTo.Text = GetRCList(medMain.MAPIMess, vbRecipTypeTo)
'    txtcc.Text = GetRCList(medMain.MAPIMess, vbRecipTypeCc)
'End Sub
'
'Sub PrintLongText(ByVal LongText As String)
''이 프로시저는 프린터에 텍스트 스트림을 인쇄하고 줄 사이에서 단어가 잘리지 않도록
''필요한 만큼 단어를 자릅니다.
'
'    Do Until LongText = ""
'        Word$ = Token$(LongText, " ")
'        If Printer.TextWidth(Word$) + Printer.CurrentX > Printer.Width - Printer.TextWidth("ZZZZZZZZ") Then
'            Printer.Print
'        End If
'        Printer.Print " " + Word$;
'    Loop
'
'End Sub
'
'Function Token$(tmp$, search$)
'    X = InStr(1, tmp$, search$)
'    If X Then
'       Token$ = Mid$(tmp$, 1, X - 1)
'       tmp$ = Mid$(tmp$, X + 1)
'    Else
'       Token$ = tmp$
'       tmp$ = ""
'    End If
'End Function
'
'Function P(ByVal expression As String, ByVal Delim As String, _
'           ByVal Piece As Integer) As String
'Dim CNTA, CNTB As Integer
'    P = ""
'    CNTA = 0
'    CNTB = 0
'    Do
'        If CNTB = Piece - 1 Then Exit Do
'        CNTA = InStr(CNTA + 1, expression, Delim)
'        If CNTA <> 0 Then CNTB = CNTB + 1
'    Loop Until CNTA = 0
'    If CNTA = 0 And Piece <> 1 Then Exit Function
'    CNTB = InStr(CNTA + 1, expression, Delim)
'    If CNTB = 0 Then CNTB = Len(expression) + 1
'    P = Mid$(expression, CNTA + 1, CNTB - CNTA - 1)
'End Function
'
'
