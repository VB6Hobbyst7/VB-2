Attribute VB_Name = "modPrintItem"
Option Explicit

Private Const FONT_KODIK        As String = "���ü"
Private Const FONT_HYKODIK      As String = "HY�߰��"
Private Const FONT_GULLIM       As String = "����ü"
Private Const FONT_BATANGCHE    As String = "����ü"
Private Const FONT_BATANG       As String = "����"
Private Const FONT_MOUNGJO      As String = "����ü"
Private Const FONT_TODUM        As String = "����ü"
Private Const FONT_TNEWROMAN    As String = "Times New Roman"
Private Const FONT_SERIF        As String = "MS Serif"

Private Const TWIP_MM           As Single = 56.7
Private Const TWIP_CM           As Single = 567

Private Const TOP_MARGIN        As Single = 500
Private Const LEFT_MARGIN       As Single = 10
Private Const BATTON_MARGIN     As Single = 567

Private Const FORM_WIDTH        As Single = TWIP_CM * 19 'A4
Private Const FORM_HEIGHT       As Single = TWIP_CM * 29 'A4
Private Const PAGE_ROW_TOT      As Single = 48 'A4

Private Const HEAD_LINE1        As Single = TOP_MARGIN + (TWIP_CM * 1)
Private Const HEAD_TEXTY1       As Single = TOP_MARGIN + (TWIP_CM * 1.2)
Private Const HEAD_LINE2        As Single = TOP_MARGIN + (TWIP_CM * 1.7)

Private Const DATA_TEXT         As Single = TOP_MARGIN + (TWIP_CM * 2)
Private Const DATA_GAP          As Single = TWIP_CM * 0.5

Private Const TAIL_LINE1        As Single = FORM_HEIGHT - (TWIP_CM * 2)
Private Const TAIL_TEXTY1       As Single = FORM_HEIGHT - (TWIP_CM * 1.8)

Private Const HEAD_TEXTX1       As Single = LEFT_MARGIN + (TWIP_CM * 0.1)
Private Const HEAD_TEXTX2       As Single = LEFT_MARGIN + (TWIP_CM * 2.5)
Private Const HEAD_TEXTX3       As Single = LEFT_MARGIN + (TWIP_CM * 5.5)
Private Const HEAD_TEXTX4       As Single = LEFT_MARGIN + (TWIP_CM * 10)
Private Const HEAD_TEXTX5       As Single = LEFT_MARGIN + (TWIP_CM * 13.5)

Public Sub PrintFrom(prtItems As ListItems)
    Call Header
    Call Body(prtItems)
    Call Tail
    Printer.EndDoc
End Sub

Private Sub Body(prtItems As ListItems)
    Dim lngCnt  As Long
    Dim itemX   As ListItem
    
    For Each itemX In prtItems
        Call PrintText(HEAD_TEXTX1, DATA_TEXT + (DATA_GAP * lngCnt), itemX.Index, , 9)
        Call PrintText(HEAD_TEXTX2, DATA_TEXT + (DATA_GAP * lngCnt), itemX.Text, , 9, True)
        Call PrintText(HEAD_TEXTX3, DATA_TEXT + (DATA_GAP * lngCnt), Trim(itemX.SubItems(1)), , 9)
        Call PrintText(HEAD_TEXTX4, DATA_TEXT + (DATA_GAP * lngCnt), Trim(itemX.SubItems(2)), , 9)
        Call PrintText(HEAD_TEXTX5, DATA_TEXT + (DATA_GAP * lngCnt), Trim(itemX.SubItems(3)), , 9)
        lngCnt = lngCnt + 1
        
        If (lngCnt Mod PAGE_ROW_TOT) = 0 Then
            Call Tail
            Printer.NewPage
            Call Header
            lngCnt = 0
        End If
    Next
    
End Sub

Private Sub Header()
    Dim strTitle As String
    strTitle = INS_NAME & " TEST CODE"
    Printer.Font = FONT_TNEWROMAN
    Printer.FontSize = 18
    Printer.Fontbold = True
    Call PrintText((FORM_WIDTH / 2) - (Printer.TextWidth(strTitle) / 2), TOP_MARGIN, strTitle, FONT_TNEWROMAN, 18, True)
    
    Printer.Line (LEFT_MARGIN, HEAD_LINE1)-(FORM_WIDTH, HEAD_LINE1 + 5), , BF
    Call PrintText(HEAD_TEXTX1, HEAD_TEXTY1, "����")
    Call PrintText(HEAD_TEXTX2, HEAD_TEXTY1, "��� �ڵ�")
    Call PrintText(HEAD_TEXTX3, HEAD_TEXTY1, "��� �˻��")
    Call PrintText(HEAD_TEXTX4, HEAD_TEXTY1, "LIS �ڵ�")
    Call PrintText(HEAD_TEXTX5, HEAD_TEXTY1, "LIS �˻��")
    Printer.Line (LEFT_MARGIN, HEAD_LINE2)-(FORM_WIDTH, HEAD_LINE2 + 5), , BF
End Sub

Private Sub Tail()
    Printer.Line (LEFT_MARGIN, TAIL_LINE1)-(FORM_WIDTH, TAIL_LINE1 + 5), , BF
    Call PrintText(LEFT_MARGIN, TAIL_TEXTY1, "����� :" & Format(Now, "YYYY�� MM�� DD��"), , , True)
    Call PrintText(HEAD_TEXTX5, TAIL_TEXTY1, "INJE UNIVERSITY PAIK HOSPITAL", , , True)
End Sub

Private Sub PrintText(ByVal x As Single, ByVal y As Single, ByVal prtText As String, _
                        Optional ByVal strFont As String = FONT_BATANGCHE, _
                        Optional ByVal FontSize As Long = 10, _
                        Optional ByVal Fontbold As Boolean = False)
   
   Dim oldFontSize As Integer
   
   Printer.CurrentX = x 'FONT_WIDTH * X
   Printer.CurrentY = y 'FONT_HEIGHT * Y
   Printer.FontSize = FontSize
   Printer.FontName = strFont
   Printer.Fontbold = Fontbold
   
   Printer.Print prtText
   
End Sub
