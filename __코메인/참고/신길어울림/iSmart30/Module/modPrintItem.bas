Attribute VB_Name = "modPrintItem"
Option Explicit

Private Const FONT_KODIK        As String = "고딕체"
Private Const FONT_HYKODIK      As String = "HY견고딕"
Private Const FONT_GULLIM       As String = "굴림체"
Private Const FONT_BATANGCHE    As String = "바탕체"
Private Const FONT_BATANG       As String = "바탕"
Private Const FONT_MOUNGJO      As String = "명조체"
Private Const FONT_TODUM        As String = "돋움체"
Private Const FONT_TNEWROMAN    As String = "Times New Roman"
Private Const FONT_SERIF        As String = "MS Serif"

Private Const TWIP_MM           As Single = 56.7
Private Const TWIP_CM           As Single = 567

Private Const TOP_MARGIN        As Single = 1000
Private Const LEFT_MARGIN       As Single = 10
Private Const BATTON_MARGIN     As Single = 567

Private Const FORM_WIDTH        As Single = TWIP_CM * 19 'A4
Private Const FORM_HEIGHT       As Single = TWIP_CM * 29 'A4
Private Const PAGE_ROW_TOT      As Single = 35 'A4

Private Const HEAD_LINE1        As Single = TOP_MARGIN + (TWIP_CM * 1)
Private Const HEAD_TEXTY1       As Single = TOP_MARGIN + (TWIP_CM * 1.2)
Private Const HEAD_LINE2        As Single = TOP_MARGIN + (TWIP_CM * 1.7)

Private Const DATA_TEXT         As Single = TOP_MARGIN + (TWIP_CM * 2)
Private Const DATA_GAP          As Single = TWIP_CM * 0.5

Private Const TAIL_LINE1        As Single = FORM_HEIGHT - (TWIP_CM * 2)
Private Const TAIL_TEXTY1       As Single = FORM_HEIGHT - (TWIP_CM * 1.8)

Private Const HEAD_TEXTX1       As Single = LEFT_MARGIN + (TWIP_CM * 1)
Private Const HEAD_TEXTX2       As Single = LEFT_MARGIN + (TWIP_CM * 3)
Private Const HEAD_TEXTX3       As Single = LEFT_MARGIN + (TWIP_CM * 5)
Private Const HEAD_TEXTX4       As Single = LEFT_MARGIN + (TWIP_CM * 7)
Private Const HEAD_TEXTX5       As Single = LEFT_MARGIN + (TWIP_CM * 8.5)
Private Const HEAD_TEXTX6       As Single = LEFT_MARGIN + (TWIP_CM * 10.1)
Private Const HEAD_TEXTX7       As Single = LEFT_MARGIN + (TWIP_CM * 13.3)

Public Sub PrintFrom(ByVal brspread As Object)
    Call Header
    Call Body(brspread)
    Call Tail
    Printer.EndDoc
End Sub

Private Sub Body(ByVal brspread1 As Object)
    Dim lngCnt  As Long
    Dim yRow  As Integer, yCol As Integer
    Dim Printtmp As Variant
    Dim Codalist As String
    
    With brspread1
        For yRow = 1 To .MaxRows
            
            .GetText 0, yRow, Printtmp
            Call PrintText(HEAD_TEXTX1 + 50, DATA_TEXT + (DATA_GAP * lngCnt), Trim(Printtmp), , 9)
            .GetText 4, yRow, Printtmp
            Call PrintText(HEAD_TEXTX2 + 50, DATA_TEXT + (DATA_GAP * lngCnt), Trim(Printtmp), , 9)
            .GetText 3, yRow, Printtmp
            Call PrintText(HEAD_TEXTX3 + 50, DATA_TEXT + (DATA_GAP * lngCnt), Trim(Printtmp), , 9)
            .GetText 6, yRow, Printtmp
            Call PrintText(HEAD_TEXTX4 + 50, DATA_TEXT + (DATA_GAP * lngCnt), Trim(Printtmp), , 9)
            .GetText 7, yRow, Printtmp
            Call PrintText(HEAD_TEXTX5 + 50, DATA_TEXT + (DATA_GAP * lngCnt), Trim(Printtmp), , 9)
            
            Codalist = "/"
            For yCol = 8 To .MaxCols
                .Col = yCol: .Row = yRow
                If .BackColor = &H80FF80 Then
                    .GetText yCol, 0, Printtmp
                    Codalist = Codalist & "/" & Trim(Printtmp)
                    Codalist = Replace(Codalist, "//", "")
                    
                    If Len(Codalist) > 50 And yCol <> .MaxCols Then
                        Codalist = Codalist & "/..."
                        Exit For
                    End If
                End If
            Next yCol
            
            Call PrintText(HEAD_TEXTX6, DATA_TEXT + (DATA_GAP * lngCnt), Trim(Codalist), , 9)
    
            lngCnt = lngCnt + 1
            
            If (lngCnt Mod PAGE_ROW_TOT) = 0 Then
                Call Tail
                Printer.NewPage
                Call Header
                lngCnt = 0
            End If
        Next
    End With
    
End Sub

Private Sub Header()
    Dim strTitle As String
    strTitle = INS_NAME & " WorkList"
    Printer.Font = FONT_TNEWROMAN
    Printer.FontSize = 18
    Printer.Fontbold = True
    Call PrintText((FORM_WIDTH / 2) - (Printer.TextWidth(strTitle) / 2) + 300, TOP_MARGIN, strTitle, FONT_TNEWROMAN, 18, True)
    
    Printer.Line (LEFT_MARGIN + 300, HEAD_LINE1)-(FORM_WIDTH + 300, HEAD_LINE1 + 10), , BF
    Call PrintText(HEAD_TEXTX1, HEAD_TEXTY1, "순서", , , True)
    Call PrintText(HEAD_TEXTX2, HEAD_TEXTY1, "환자명", , , True)
    Call PrintText(HEAD_TEXTX3, HEAD_TEXTY1, "검체번호", , , True)
    Call PrintText(HEAD_TEXTX4, HEAD_TEXTY1, "RackNo", , , True)
    Call PrintText(HEAD_TEXTX5, HEAD_TEXTY1, "PosNo", , , True)
    Call PrintText(HEAD_TEXTX7, HEAD_TEXTY1, "검사항목", , , True)
    Printer.Line (LEFT_MARGIN + 300, HEAD_LINE2)-(FORM_WIDTH + 300, HEAD_LINE2 + 10), , BF
End Sub

Private Sub Tail()
    Printer.Line (LEFT_MARGIN + 300, TAIL_LINE1)-(FORM_WIDTH + 300, TAIL_LINE1 + 10), , BF
    Call PrintText(LEFT_MARGIN + 330, TAIL_TEXTY1, HOS_NAME, , , True)
'    Call PrintText(HEAD_TEXTX5, TAIL_TEXTY1, HOS_NAME, , , True)
End Sub

Private Sub PrintText(ByVal X As Single, ByVal Y As Single, ByVal prtText As String, _
                        Optional ByVal strFont As String = FONT_BATANGCHE, _
                        Optional ByVal FontSize As Long = 10, _
                        Optional ByVal Fontbold As Boolean = False)
   
   Dim oldFontSize As Integer
   
   Printer.CurrentX = X 'FONT_WIDTH * X
   Printer.CurrentY = Y 'FONT_HEIGHT * Y
   Printer.FontSize = FontSize
   Printer.FontName = strFont
   Printer.Fontbold = Fontbold
   
   Printer.Print prtText
   
End Sub
