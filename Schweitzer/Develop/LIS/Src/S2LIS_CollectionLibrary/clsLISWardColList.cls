VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISWardColList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private iPWACoHeRow As Integer
Private fWhich As Object
Private iPageWidth As Integer
Private iPageHeight As Integer
Private iCurY As Integer
Private DataExist As Boolean
Private sLastDt As String
Private sLastTm As String
Private iRecordCount As Integer

Private pWardId As String
Private pWardNm As String
Private pWorkDt As String
Private pWorkTm As String
Private pBuildCd As String
Private pBuildNm As String
    
Private SvBuildCd As String
Private SvBuildNm As String

Const iCm = 567
Const iLineHeight = 10

Private iposSEQ%, iposWorkNo%, iposPtName%, iposPtID%, iposSAge%, _
    iposIO%, iposRcv, iposSF%, iposTestCD%, iposSpccd%


Public Property Get WardId() As String
    WardId = pWardId
End Property

Public Property Let WardId(ByVal vNewValue As String)
    pWardId = vNewValue
End Property

Public Property Get WardNm() As String
    WardNm = pWardNm
End Property

Public Property Let WardNm(ByVal vNewValue As String)
    pWardNm = vNewValue
End Property

Public Property Get WorkDt() As String
    WorkDt = pWorkDt
End Property

Public Property Let WorkDt(ByVal vNewValue As String)
    pWorkDt = vNewValue
End Property

Public Property Get WorkTm() As String
    WorkTm = pWorkTm
End Property

Public Property Let WorkTm(ByVal vNewValue As String)
    pWorkTm = vNewValue
End Property

Public Property Get BuildCd() As String
    BuildCd = pBuildCd
End Property

Public Property Let BuildCd(ByVal vNewValue As String)
    pBuildCd = vNewValue
End Property

Public Property Get BuildNm() As String
    BuildNm = pBuildNm
End Property

Public Property Let BuildNm(ByVal vNewValue As String)
    pBuildNm = vNewValue
End Property


Public Sub Print_ColList()

   Dim SqlStmt As String
    
   Call PrtBody
    
   Dim tmpRs As DrRecordSet
   
   SqlStmt = "select field1 as ColDate, field2 as ColTime from " & TB_LAB031 & _
                " where cdindex = '" & CD1_ColListTm & "' and cdval1 = '" & pBuildCd & "' " & _
                " and     cdval2 = '" & pWardId & "' "
   Set tmpRs = OpenRecordSet(SqlStmt)
   
   If tmpRs.EOF Then
      DataExist = False
   Else
      DataExist = True
   End If
   
   tmpRs.RsClose
   Set tmpRs = Nothing
    
    If DataExist Then
      SqlStmt = "update " & TB_LAB031 & " set field1 = '" & pWorkDt & "', " & _
                                                        "      field2 = '" & sLastTm & "' " & _
                                                        " where cdindex = '" & CD1_ColListTm & "' " & _
                                                        " and     cdval1 = '" & pBuildCd & "' " & _
                                                        " and     cdval2 = '" & pWardId & "' "
    Else
      SqlStmt = "insert into " & TB_LAB031 & "(cdindex, cdval1, cdval2, field1, field2) " & _
                    "values ('" & CD1_ColListTm & "', '" & pBuildCd & "', '" & pWardId & "', " & _
                    "            '" & pWorkDt & "', '" & sLastTm & "') "
    End If
    
    dbconn.BeginTrans
    dbconn.Execute (SqlStmt)
    dbconn.CommitTrans
    
End Sub

Private Sub InitReport()
    iPageWidth = Printer.ScaleWidth
    iPageHeight = Printer.ScaleHeight
End Sub

Private Sub PrtHeader()
   
    Dim TITLE As String
    
    '/* 보고서 제목
    TITLE = "채혈 리스트"
    
    
    '        "채혈 리스트"
    ' ----------------------
    ' ----------------------
    
    Call prtTitle(TITLE, iCm / 4)
    Call DrawLine(8 * iCm, iCurY, 13 * iCm, iCurY, "dot", 1, iCm / 10)
    Call DrawLine(8 * iCm, iCurY, 13 * iCm, iCurY, "dot", 1, iCm * 1.5)
    
    ' -----------------------------------------------------------------------------
    Call DrawLine(iCm / 2, iCurY, iPageWidth - iCm / 2, iCurY, "solid", 2, iCm / 4)
    
    iposSEQ = iCm / 2
    iposWorkNo = iposSEQ + iCm - iCm / 4
    iposPtName = iposWorkNo + 2.3 * iCm
    iposPtID = iposPtName + iCm + iCm / 2
    iposSAge = iposPtID + iCm + iCm / 2
    iposIO = iposSAge + iCm
    iposRcv = iposIO + iCm
    iposSF = iposRcv + 2.5 * iCm
    iposTestCD = iposSF + 0.8 * iCm
    iposSpccd = iposTestCD + 7.8 * iCm
    
    Call WriteStr(iCurY, iposSEQ, "SEQ", iCurY, 0)
    Call WriteStr(iCurY, iposWorkNo, "    Work No", iCurY, 0)
    Call WriteStr(iCurY, iposPtName, "환자성명", iCurY, 0)
    Call WriteStr(iCurY, iposPtID, "  환자ID", iCurY, 0)
    Call WriteStr(iCurY, iposSAge, "S/Age", iCurY, 0)
    Call WriteStr(iCurY, iposIO, "  I/O", iCurY, 0)
    Call WriteStr(iCurY, iposRcv, "      채혈일시", iCurY, 0)
    Call WriteStr(iCurY, iposSF, " S/F", iCurY, 0)
    Call WriteStr(iCurY, iposTestCD, "                              검사종목", iCurY, 0)
    Call WriteStr(iCurY, iposSpccd, "  검체", iCurY, iCm / 2)
    
    Call DrawLine(iCm / 2, iCurY, iPageWidth - iCm / 2, iCurY, "solid", 2, 0)
    
 
End Sub

Private Sub WriteStr(ByVal Y As Integer, ByVal X As Integer, ByVal str As String, _
                    iNextY As Integer, iSpace As Integer)
    Printer.CurrentY = Y
    Printer.CurrentX = X
    iNextY = Printer.CurrentY + iSpace
    Printer.Print str

   
End Sub


Private Sub prtTitle(TITLE As String, iSpace As Integer)

    Dim oldFontSize As Integer
    
    oldFontSize = Printer.FontSize
    Printer.FontSize = 14
    Printer.FontBold = True
    '/* Tile이 중앙으로 오도록 string길이에 따라 위치를 계산한다.
    
    Printer.CurrentY = 0
    Printer.CurrentX = iPageWidth / 2 - Printer.TextWidth(TITLE) / 2
    iCurY = Printer.CurrentY + Printer.TextHeight(TITLE) + iSpace
    
    Printer.Print TITLE
    Printer.FontSize = oldFontSize
    Printer.FontBold = False

End Sub

Private Sub ChangeLine(iLineSpace As Integer)
    iCurY = iCurY + iLineSpace
    Printer.CurrentY = iCurY
    Printer.CurrentX = iCm / 2
    
End Sub

Private Sub PrtBody()

    Dim sSQL1 As String
    Dim sSQL2 As String
    Dim rsWorksheet As Object
    Dim rsPtName As Object
    
    Dim tmpStr As String, sStsCd As String
    Dim I%
    
    'If optBusiDiv(0).Value Then
    '     sStsCd = STS_Access    '외래-->접수상태
    'Else
         sStsCd = STS_HaveSpc    '병동-->채혈상태
    'End If

    Printer.FontSize = 9
    
    sSQL1 = " select a.workarea, a.accdt, a.accseq, a.ptid, a.sex, a.ageday, a.deptcd, a.coldt, " & _
                " a.coltm, a.storecd, a.spccd, a.testdiv, a.buildcd, b.field1 as BuildNm, d.field3 as SpcNm " & _
                " from " & TB_LAB201 & " a, " & TB_LAB032 & " b, " & TB_LAB204 & " c, " & TB_LAB032 & " d " & _
                " where c.workdt = '" & pWorkDt & "' and " & _
                "       c.wardid = '" & pWardId & "' and " & _
                "       c.worktm = '" & pWorkTm & "' and " & _
                "       c.buildcd = '" & pBuildCd & "' and " & _
                "       a.workarea = c.workarea  and " & _
                "       a.accdt  = c.accdt  and " & _
                "       a.accseq = c.accseq  and " & _
                "       b.cdindex = '" & CD2_Buildings & "' and " & _
                "       b.cdval1 = a.buildcd and " & _
                "       d.cdindex = '" & CD2_Specimen & "' and " & _
                "       d.cdval1 = a.spccd " & _
                " order by a.buildcd, a.workarea, a.accdt, a.accseq "

    Set rsWorksheet = OpenRecordSet(sSQL1)
   
    If rsWorksheet.EOF = True Then ' record가 존재하지 않을경우
       MsgBox " 데이타가 존재하지 않습니다. "
       GoTo NoData
    End If
    
    iRecordCount = rsWorksheet.RecordCount
    
    SvBuildCd = Trim(rsWorksheet.Fields("buildcd").Value)
    SvBuildNm = Trim(rsWorksheet.Fields("buildnm").Value)
    
    Call InitReport
    Call PrtHeader
    Call prtPageNum
    Call prtTerm

    Dim temp1 As String, temp2 As String
    Dim sAge As String
    Dim iSeqNum As Integer
    
    sLastTm = ""
    With rsWorksheet
    
        
        If sLastTm < .Fields("ColTm").Value Then sLastTm = .Fields("ColTm").Value
            
        temp1 = Mid(.Fields("ColTm").Value, 1, 4)
        temp2 = Format(temp1, "00:00")
        .MoveFirst
        
        For I = 1 To .RecordCount
            
            sSQL2 = " select ptnt_nm as ptnm from " & TB_HIS001 & _
                    " where ptnt_no = " & rsWorksheet.Fields("ptid").Value
            Set rsPtName = OpenRecordSet(sSQL2)

            If chkTestCD(.Fields("WorkArea").Value, .Fields("accdt").Value, .Fields("accseq").Value, _
                         .Fields("testdiv").Value) = True Then                   ' Exists
                Call ChangeLine(iCm / 2)
                
               If SvBuildCd <> Trim(.Fields("buildcd").Value) Then
                    SvBuildCd = Trim(.Fields("buildcd").Value)
                    SvBuildNm = Trim(.Fields("buildnm").Value)
                    Printer.NewPage
                    Call PrtHeader
                    Call prtPageNum
                    Call prtTerm
                    Call ChangeLine(iCm / 2)
                  'Call WriteStr(iCurY, iposSEQ + iCm / 6, CStr(iSeqNum), iCurY, 0)
               End If
                
                If iCurY > iPageHeight - 2 * iCm Then ' newPage일 경우
                    Printer.NewPage
                    Call PrtHeader
                    Call prtPageNum
                    Call prtTerm
                    Call ChangeLine(iCm / 2)
                End If
            
                iSeqNum = iSeqNum + 1
            
                sAge = (.Fields("AgeDay").Value \ 365) + 1
                Call WriteStr(iCurY, iposSEQ + iCm / 6, CStr(iSeqNum), iCurY, 0)
                Call WriteStr(iCurY, iposWorkNo + iCm / 6, Trim(CStr(.Fields("WorkArea").Value)) & "-" & Del2Chr(.Fields("AccDt").Value) & _
                              "-" & Trim(CStr(.Fields("AccSeq").Value)), iCurY, 0)
                Call WriteStr(iCurY, iposPtName + iCm / 6, rsPtName.Fields("ptnm").Value, iCurY, 0)

                Call WriteStr(iCurY, iposPtID + iCm / 6, Trim(CStr(.Fields("PtId").Value)), iCurY, 0)
                Call WriteStr(iCurY, iposSAge + iCm / 6, Trim(.Fields("Sex").Value) & "/" & Trim(CStr(sAge)), iCurY, 0)
                Call WriteStr(iCurY, iposIO + iCm / 6, Trim(.Fields("DeptCd").Value), iCurY, 0)
                Call WriteStr(iCurY, iposRcv + iCm / 6, Del2Chr(.Fields("ColDt").Value) & _
                              "   " & temp2, iCurY, 0)
                Call WriteStr(iCurY, iposSF + iCm / 6, Trim(.Fields("StoreCd").Value), iCurY, 0)
                Call WriteStr(iCurY, iposSpccd + iCm / 6, Trim(.Fields("SpcNm").Value), iCurY, 0)
            
                Call WriteTestCD(.Fields("WorkArea").Value, .Fields("AccDt").Value, .Fields("AccSeq").Value, .Fields("TestDiv").Value)
            End If
            .MoveNext
        Next I
        rsPtName.RsClose
        Set rsPtName = Nothing
    End With
        
    Printer.EndDoc
    
NoData:
    rsWorksheet.RsClose
    Set rsWorksheet = Nothing
    
End Sub

Private Function Del2Chr(sStr As String) As String
    Del2Chr = Trim(Mid(sStr, 3, Len(sStr) - 2))
End Function

Private Function chkTestCD(sWorkarea As String, sAccdt As String, sAccSeq As String, stestdiv As String) As Boolean
    Dim sSQL2 As String
    Dim rsTestCode As Object
            
    sSQL2 = " select ordcd " & _
        " from " & TB_LAB102 & _
        " where workarea = '" & sWorkarea & "' and " & _
        " accdt = '" & Trim(sAccdt) & "' and " & _
        " accseq = " & Trim(sAccSeq)
    
    Set rsTestCode = OpenRecordSet(sSQL2)
    
    If rsTestCode.EOF = True Then
        chkTestCD = False         ' not exitst
    Else
        chkTestCD = True              ' Exist
    End If
    
    rsTestCode.RsClose
    Set rsTestCode = Nothing
    
End Function

Private Sub WriteTestCD(sWorkarea As String, sAccdt As String, sAccSeq As String, stestdiv As String)
    Dim sSQL2 As String
    Dim rsTestCode As Object
    Dim I%, tmpiposTestCD
    Dim sTable As String
    
    sSQL2 = " select ordcd " & _
        " from " & TB_LAB102 & _
        " where workarea = '" & sWorkarea & "' and " & _
        " accdt = '" & Trim(sAccdt) & "' and " & _
        " accseq = " & Trim(sAccSeq)
    
    Set rsTestCode = OpenRecordSet(sSQL2)
    
    If rsTestCode.EOF = True Then
'        MsgBox " 레코드가 없다"
        GoTo NoData
    End If
    
    With rsTestCode
        tmpiposTestCD = iposTestCD
        rsTestCode.MoveFirst
        For I = 1 To rsTestCode.RecordCount

            Call WriteStr(iCurY, tmpiposTestCD + iCm / 6, Trim(.Fields("OrdCd").Value), iCurY, 0)
            tmpiposTestCD = tmpiposTestCD + 1.5 * iCm
            If (I Mod 5 = 0) Then
                Call ChangeLine(iCm / 2)
                If iCurY > iPageHeight - 2 * iCm Then  ' newPage일 경우
                    Printer.NewPage
                    Call PrtHeader
                    Call prtPageNum
                    Call prtTerm
                    Call ChangeLine(iCm / 2)
                End If
                tmpiposTestCD = iposTestCD
            End If
            rsTestCode.MoveNext
        Next I
        If (rsTestCode.RecordCount Mod 5) = 0 Then
            iCurY = iCurY - iCm / 2
        End If
    End With
    
NoData:
    rsTestCode.RsClose
    Set rsTestCode = Nothing
    
End Sub

Private Sub DrawLine(ByVal iStartX As Integer, ByVal iStartY As Integer, _
                    ByVal iEndX As Integer, ByVal iEndy As Integer, _
                    sLineStyle As String, iLinewidth As Integer, iSpace As Integer)

    Select Case sLineStyle
        Case "solid"
            Printer.DrawStyle = 0
        Case "dash"
            Printer.DrawStyle = 1
        Case "dot"
            Printer.DrawStyle = 2
        Case "dashdot"
            Printer.DrawStyle = 3
        Case "dashdotdot"
            Printer.DrawStyle = 4
    End Select
         
    Printer.DrawWidth = iLinewidth
    Printer.Line (iStartX, iStartY)-(iEndX, iEndy)
    iCurY = Printer.CurrentY + iSpace
End Sub

Private Sub prtPageNum()
    
    Dim oldX As Integer, oldY As Integer
    Dim sDate As String, sTime As String
    
    sDate = Format(Now, "YYYY/MM/DD")
    sTime = Format(Now, "HH:MM:SS")
    oldX = Printer.CurrentX
    oldY = Printer.CurrentY
    
    Printer.CurrentX = iPageWidth - 4 * iCm
    Printer.CurrentY = 0
    Printer.Print "P A G E  : " & Printer.Page
            
    Printer.CurrentX = iPageWidth - 4 * iCm
    Printer.CurrentY = Printer.TextHeight("P A G E") + iCm / 6
    Printer.Print "RUN-DATE : " & sDate
        
    Printer.CurrentX = iPageWidth - 4 * iCm
    Printer.CurrentY = Printer.TextHeight("P A G E") + iCm / 6 + _
                           Printer.TextHeight("RUN-DATE") + iCm / 6
    Printer.Print "RUN-TIME : " & sTime
        
    Printer.CurrentX = oldX
    Printer.CurrentY = oldY
    
End Sub

Private Sub prtTerm()
    Dim oldX As Integer, oldY As Integer
    Dim oldFontSize As Integer
    
    
    oldX = Printer.CurrentX
    oldY = Printer.CurrentY
        
    oldFontSize = Printer.FontSize
    Printer.FontSize = 11
    Printer.FontBold = True
    Printer.CurrentX = iCm
    Printer.CurrentY = 1.3 * iCm
    Printer.Print "Delivery Location  : " & SvBuildCd & "   " & SvBuildNm
    Printer.CurrentX = iCm * 8
    Printer.CurrentY = 1.3 * iCm
    'If optBusiDiv(0).Value Then
    '     Printer.Print "( 외래검체 - 접수상태 )"
    'Else
         Printer.Print "( 병동검체 - 채혈상태 )"
    'End If
    
    
    Printer.FontSize = oldFontSize
    Printer.FontBold = False
    Printer.CurrentX = iCm
    Printer.CurrentY = 1.3 * iCm + Printer.TextHeight("Work Area : ") + iCm / 4
    Printer.Print "채혈장소   : " & pWardId & "    " & pWardNm & "    /    " & _
                      "채혈일시   : " & Format(pWorkDt, CS_DateMask) & "    " & Format(pWorkTm, CS_TimeLMask)
    Printer.CurrentX = iPageWidth - 4 * iCm
    Printer.CurrentY = 1.3 * iCm + Printer.TextHeight("Work Area : ") + iCm / 4
    Printer.FontBold = True
    Printer.Print "검체수   :   총  " & CStr(iRecordCount) & "  개"
    Printer.FontBold = False
                                    
    Printer.CurrentX = oldX
    Printer.CurrentY = oldY
End Sub




