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
'Public Const conUnreadMessage = "*"     ' ���� ���� �޽����� ��Ÿ���� ���ڿ��� ���� ���
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
'' �� ���ν����� MAPI ��¥�� ���� �޽����� ���� ���� �� ���� �� �ϳ��� ���մϴ�.
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
'    Header = Header + "��: " + MSG.MsgOrigDisplayName + CR
'    Header = Header + "�޴���: " + GetRCList(MSG, vbRecipTypeTo) + CR
'    Header = Header + "����: " + GetRCList(MSG, vbRecipTypeCc) + CR
'    Header = Header + "����: " + MSG.MsgSubject + CR
'    Header = Header + "���� ��¥: " + DateFROMMapiDate$(MSG.MsgDateReceived, conMailLongDate) + CR + CR
'
'    GetHeader = Header
'
'End Function
'
'Function GetRCList(MSG As Control, RCType As Integer) As String
''�޴� ����� ����� �ִ� ��� �� �Լ��� ���� �������� ������
''�������� �޴� ����� ����� ��ȯ�մϴ�. (��� 1;��� 2;��� 3)
'
'    For i = 0 To MSG.RecipCount - 1
'        MSG.RecipIndex = i
'        If RCType = MSG.RecipType Then
'            A$ = A$ + ";" + MSG.RecipDisplayName
'        End If
'    Next i
'
'    If A$ <> "" Then
'        A$ = Mid$(A$, 2)  '���� ";"�� ����ϴ�.
'    End If
'
'    GetRCList = A$
'
'End Function
'
'Sub UpdateRecips()
'' �� ���ν����� �ùٸ� ���� �ʵ�� �޴� ��� ������ ���� ��Ĩ�ϴ�.
'    txtTo.Text = GetRCList(medMain.MAPIMess, vbRecipTypeTo)
'    txtcc.Text = GetRCList(medMain.MAPIMess, vbRecipTypeCc)
'End Sub
'
'Sub PrintLongText(ByVal LongText As String)
''�� ���ν����� �����Ϳ� �ؽ�Ʈ ��Ʈ���� �μ��ϰ� �� ���̿��� �ܾ �߸��� �ʵ���
''�ʿ��� ��ŭ �ܾ �ڸ��ϴ�.
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
