VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private lngStartX        As Long     ' Start X Position
Private lngSpace         As Long     ' 라인사이의 Size
Private lngLastYpos      As Long     ' Last Y Position
Private lngScaleWidth    As Long     ' Printer.ScaleWidth
Private lngCurYPos       As Long     ' Current Y Position
Private mvarCurYPos      As Long

Public PrinterHeader1      As String   'PrinterHeader1 큰클자
Public PrinterHeader2      As String   'PrinterHeader2 (출력자,출력일시등)
Public PrinterHeader3      As String   'PrinterHeader3 (출력값 이름)
Public PrinterBody         As String
Public PrintBoxBln         As Boolean
Public lngCount            As Long

Public Sub CallPrint(Optional ByVal pMode As String = "", Optional ByVal mvarStartX As Long = 5, _
                     Optional ByVal mvarSpace As Long = 6, Optional ByVal CurYPos As Long = 10)
    mvarCurYPos = CurYPos

    lngCount = 1
    Call PrintIntionlize(mvarStartX, mvarSpace, mvarCurYPos, pMode)
    Call BodyPrint
    
    Printer.EndDoc
End Sub

Private Sub HeaderPrint()
    Dim aryHeader2()  As String
    Dim aryHeader3()  As String
    Dim strHead       As String
    Dim strChangeLine As String
    Dim lngXPos       As Long
    Dim ii            As Long
    
    lngCurYPos = mvarCurYPos

    If lngCount < 2 Then
        If PrinterHeader1 <> "" Then
            Printer.FontSize = 20: Printer.FontBold = True
            
            Call strPrint(PrinterHeader1, lngStartX, lngSpace, lngScaleWidth, "C", "C", True)
            Printer.FontSize = 9: Printer.FontBold = False
        End If
        lngCurYPos = lngCurYPos + lngSpace
    End If
    
    
    If PrinterHeader2 <> "" Then
        aryHeader2 = Split(PrinterHeader2, vbTab)
        For ii = LBound(aryHeader2) To UBound(aryHeader2)
            strHead = medGetP(aryHeader2(ii), 1, COL_DIV)
            lngXPos = medGetP(aryHeader2(ii), 2, COL_DIV)
            strChangeLine = medGetP(aryHeader2(ii), 3, COL_DIV)
            Call strPrint(strHead, lngXPos, lngSpace, lngScaleWidth, , , IIf(strChangeLine = "1", True, False))
        Next
    End If
    Printer.DrawWidth = 7
    Printer.Line (lngStartX, lngCurYPos)-(lngScaleWidth, lngCurYPos) ': lngCurYPos = lngCurYPos + lngSpace
    
    If PrinterHeader3 <> "" Then
        aryHeader3 = Split(PrinterHeader3, vbTab)
        For ii = LBound(aryHeader3) To UBound(aryHeader3)
            strHead = medGetP(aryHeader3(ii), 1, COL_DIV)
            lngXPos = medGetP(aryHeader3(ii), 2, COL_DIV)
            strChangeLine = medGetP(aryHeader3(ii), 3, COL_DIV)
            Call strPrint(strHead, lngXPos, lngSpace, lngScaleWidth, , , IIf(strChangeLine = "1", True, False))
        Next
    End If
    
    Printer.Line (lngStartX, lngCurYPos)-(lngScaleWidth, lngCurYPos)
    
    Printer.Line (lngStartX, lngLastYpos - lngSpace)-(lngScaleWidth, lngLastYpos - lngSpace)
    
    Printer.CurrentX = lngStartX: Printer.CurrentY = lngLastYpos - lngSpace
    
    Printer.Print "예수병원" & " Page - " & lngCount
    If lngCount = 1 And PrintBoxBln = True Then Call PrintBox
    
End Sub
Private Sub BodyPrint()
    Dim aryBody()       As String
    Dim strHead         As String
    Dim strChangeLine   As String
    Dim strLine         As String
    
    Dim lngXPos         As Long
    Dim ii              As Long
    
    Call HeaderPrint
    
    If PrinterBody <> "" Then
        aryBody = Split(PrinterBody, vbTab)
        For ii = LBound(aryBody) To UBound(aryBody)
            Call CheckNewPage
            If aryBody(ii) <> "" Then
                strHead = medGetP(aryBody(ii), 1, COL_DIV)
                lngXPos = medGetP(aryBody(ii), 2, COL_DIV)
                strChangeLine = medGetP(aryBody(ii), 3, COL_DIV)
                strLine = medGetP(aryBody(ii), 4, COL_DIV)
                Call strPrint(strHead, lngXPos, lngSpace, lngScaleWidth, , , IIf(strChangeLine = "1", True, False))
    
                If strLine = "1" Then
                    Printer.DrawStyle = 1
                    Printer.DrawWidth = 1
                    Printer.Line (lngStartX, lngCurYPos)-(lngScaleWidth, lngCurYPos)
                End If
            Else
            
            End If
        Next
    End If
    
End Sub

Private Sub CheckNewPage()

    If lngCurYPos > lngLastYpos - lngSpace * 2 Then ' newPage일 경우
        Printer.NewPage
        lngCount = lngCount + 1
        Call HeaderPrint
    End If
            
End Sub


Private Sub PrintIntionlize(Optional ByVal lngXPos As Long, Optional ByVal lngLineSpace As Long, Optional ByVal lngCYpos As Long, _
                           Optional ByVal PrintScale As String = "")
    
    lngStartX = lngXPos: lngSpace = lngLineSpace: lngCurYPos = lngCYpos
    
    With Printer
        .Font = "굴림체"
        .FontSize = 9
        If PrintScale = "" Then
            .Orientation = vbPRORPortrait
        Else
            .Orientation = vbPRORLandscape
        End If
        .ScaleMode = vbMillimeters
        lngScaleWidth = .ScaleWidth
        lngLastYpos = .ScaleHeight
    End With
    
End Sub

Public Sub strPrint(ByVal sStr As String, ByVal aBaseX As Single, ByVal aBaseY As Single, _
                    Optional ByVal SpcWidth As Single, Optional ByVal WAlign As String, _
                    Optional ByVal HAlign As String, Optional ByVal blnLineAdd As Boolean = False)
    '/* 가로 정렬 */
    Select Case WAlign
        Case "C", "c"  '/* 가운데 정렬*/
            Printer.CurrentX = aBaseX + (SpcWidth - Printer.TextWidth(sStr)) / 2
        Case "R", "r"  '/* 오른쪽 정렬 */
            Printer.CurrentX = aBaseX + (SpcWidth - Printer.TextWidth(sStr)) - 0.5
        Case Else      '/* 왼쪽 정렬 */
            Printer.CurrentX = aBaseX + 0.5
    End Select
    
    '/* 세로 정렬 */
    Select Case HAlign
        Case "C", "c", "M", "m" '/* 중앙정렬 */
            Printer.CurrentY = lngCurYPos + (aBaseY - Printer.TextHeight(sStr)) / 2
'                    Printer.CurrentY = aBaseY + (SpcHeight - Printer.TextHeight(sStr)) / 2
        Case "B", "b" '/* 아래정렬 */
            Printer.CurrentY = lngCurYPos + (aBaseY - Printer.TextHeight(sStr)) - 1
        Case Else     '/* 위쪽정렬 */
            Printer.CurrentY = lngCurYPos + 1
    End Select
    
    If blnLineAdd Then lngCurYPos = lngCurYPos + aBaseY
    
    Printer.Print sStr
End Sub

Private Sub PrintBox()
    Dim lngBoxLeft  As Long
    Dim lngBoxTop   As Long
    Dim lngBoxRight As Long
    Dim lngBoxHeight As Long
    
    lngBoxTop = 5
    lngBoxLeft = lngScaleWidth - 50
    lngBoxRight = lngScaleWidth - 2
    lngBoxHeight = lngBoxTop + 17
    
    Printer.Line (lngBoxLeft, lngBoxTop)-(lngBoxRight, lngBoxHeight), , B
    Printer.Line (lngBoxLeft + 8, lngBoxTop + 5)-(lngBoxRight, lngBoxTop + 5)
    Printer.Line (lngBoxLeft + 8, lngBoxTop)-(lngBoxLeft + 8, lngBoxHeight)
    
    Printer.Line (lngBoxLeft + 21, lngBoxTop)-(lngBoxLeft + 21, lngBoxHeight)
    Printer.Line (lngBoxLeft + 35, lngBoxTop)-(lngBoxLeft + 35, lngBoxHeight)
    
    Printer.CurrentX = lngBoxLeft + 11: Printer.CurrentY = lngBoxTop + 1: Printer.Print "담 당"
    Printer.CurrentX = lngBoxLeft + 24: Printer.CurrentY = lngBoxTop + 1: Printer.Print "주 임"
    Printer.CurrentX = lngBoxLeft + 38: Printer.CurrentY = lngBoxTop + 1: Printer.Print "과 장"
    
    Printer.CurrentX = lngBoxLeft + 2: Printer.CurrentY = lngBoxTop + 2: Printer.Print "결"
    Printer.CurrentX = lngBoxLeft + 2: Printer.CurrentY = lngBoxTop + 12: Printer.Print "제"
    
End Sub

Private Sub Class_Initialize()
    PrintBoxBln = False
End Sub


