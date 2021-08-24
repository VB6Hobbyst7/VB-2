VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWorkListM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private iPageWidth As Integer
Private iPageHeight As Integer
Private iCurY As Integer
Private iPageCnt As Long
Private lngCm As Integer


Type tLabno
    sWorkarea As String
    sAccDt As String
    iAccSeq As Long
End Type

Type tLastRst
    sPtid As String
    sTestCd As String
    sSpcCd As String
End Type

Const iCm = 567
Const iLineHeight = 10

Private iposSEQ%, iposWorkNo%, iposPtName%, iposPtID%, iposSAge%, _
    iposIO%, iposRcv, iposSF%, iposTestCD%, iposSpccd%, iposAnti%

' Input Data ''''''''''''''''''
Private pSpeGroupCd As String '
Private pWsUnit As String     '
Private pSpeGroupNm As String '
Private pWorksheet2 As Boolean
Private pStainWorksheet As Boolean
'''''''''''''''''''''''''''''''
Dim pSpeGroupField4 As String

Public Property Let StainWorksheet(ByVal pValue As Boolean)
    pStainWorksheet = pValue
End Property

Public Property Let Worksheet2(ByVal pValue As Boolean)
    pWorksheet2 = pValue
End Property


Public Sub GetInputData(SpeGroupCd As String, WsUnit As String, _
                        SpeGroupNm As String)
    Dim sSqlGetSpeGroupField4 As String
    Dim rsGetSpeGroupField4 As Recordset

    pSpeGroupCd = SpeGroupCd
    pWsUnit = WsUnit
    pSpeGroupNm = SpeGroupNm

    '## 이전쿼리
    sSqlGetSpeGroupField4 = " select field1, field4 " & _
                            " from " & T_LAB032 & _
                            " where " & DBW("cdindex=", LC3_SGroup) & _
                            " and " & DBW("cdval1=", pSpeGroupCd)

    Set rsGetSpeGroupField4 = New Recordset
    rsGetSpeGroupField4.Open sSqlGetSpeGroupField4, DBConn

    pSpeGroupField4 = "" & rsGetSpeGroupField4.Fields("field4").Value
    pSpeGroupNm = "" & rsGetSpeGroupField4.Fields("field1").Value

    Set rsGetSpeGroupField4 = Nothing
End Sub

Public Sub PrintReport()
    Call InitReport
    If Not pWorksheet2 Then
        Call PrtHeader
        Call prtTerm
        Call prtPageNum
        Call Print_WaterMark
    End If
    Call PrtBody
    Call Print_WaterMark
    Printer.EndDoc
End Sub
Public Sub InitReport()
    Printer.ScaleMode = vbTwips
    iPageCnt = 0
    iPageWidth = Printer.ScaleWidth
    iPageHeight = Printer.ScaleHeight
    lngCm = CInt(iPageWidth / 20.15)
End Sub
Public Sub prtTitle(Title As String, iSpace As Integer)

    Dim oldFontSize As Integer
    Dim lngPosY  As Long

    oldFontSize = Printer.FontSize
    Printer.FontSize = 14
    Printer.FontBold = True
    '/* Tile이 중앙으로 오도록 string길이에 따라 위치를 계산한다.

    iPageCnt = iPageCnt + 1
    If pWorksheet2 And (iPageCnt Mod 2) = 0 Then
        lngPosY = iPageHeight / 2
    Else
        lngPosY = 0
    End If
    'Printer.CurrentY = 0
    
    Printer.CurrentY = lngPosY
    Printer.CurrentX = iPageWidth / 2 - Printer.TextWidth(Title) / 2

    Printer.Print Title
    Printer.FontSize = oldFontSize
    Printer.FontBold = False

    iCurY = lngPosY + Printer.TextHeight(Title) + iSpace
    Call ChangeLine(iCurY, lngCm / 10)
    Call DrawLine(iPageWidth / 2 - Printer.TextWidth(Title), iCurY, iPageWidth / 2 + Printer.TextWidth(Title), _
                  iCurY, "dot", 1)
    Call ChangeLine(iCurY, lngCm / 10)
    Call DrawLine(iPageWidth / 2 - Printer.TextWidth(Title), iCurY, iPageWidth / 2 + Printer.TextWidth(Title), _
                  iCurY, "dot", 1)
    Call ChangeLine(iCurY, lngCm / 10)

    Printer.FontSize = 11
    Printer.CurrentY = iCurY
    Printer.CurrentX = iPageWidth / 2 - Printer.TextWidth("WorkSheet Unit : " & pWsUnit) / 2

    Printer.Print "WorkSheet Unit : " & pWsUnit
    Printer.FontSize = oldFontSize
    Printer.FontBold = False
    Call ChangeLine(iCurY, lngCm)

    'iCurY =  , 9416   '1239 , 8799

End Sub

Public Sub DrawLine(ByVal iStartX As Integer, ByVal iStartY As Integer, _
                    ByVal iEndX As Integer, ByVal iEndy As Integer, _
                    sLineStyle As String, iLinewidth As Integer)

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
    'iCurY = Printer.CurrentY + iSpace
End Sub


'Public Sub DrawLine(ByVal iStartX As Integer, ByVal iStartY As Integer, _
'                    ByVal iEndX As Integer, ByVal iEndy As Integer, _
'                    sLineStyle As String, iLinewidth As Integer)
'
'    Select Case sLineStyle
'        Case "solid"
'            Printer.DrawStyle = 0
'        Case "dash"
'            Printer.DrawStyle = 1
'        Case "dot"
'            Printer.DrawStyle = 2
'        Case "dashdot"
'            Printer.DrawStyle = 3
'        Case "dashdotdot"
'            Printer.DrawStyle = 4
'    End Select
'
'    Printer.DrawWidth = iLinewidth
'    Printer.Line (iStartX, iStartY)-(iEndX, iEndy)
''    iCurY = Printer.CurrentY + iSpace
'End Sub

Public Sub WriteStr(ByVal Y As Integer, ByVal X As Integer, ByVal str As String, _
                    iNextY As Integer, iSpace As Integer)
    Printer.CurrentY = Y
    Printer.CurrentX = X
    iNextY = Printer.CurrentY + iSpace
    Printer.Print str
End Sub

Private Sub ChangeLine(iNextY As Integer, iLineSpace As Integer)

    iNextY = Printer.CurrentY + iLineSpace

End Sub

Public Sub PrtHeader()

    Dim Title As String
    Dim sWSNanme As String
    Dim iHeaderPosY As Integer

    '/* 보고서 제목
    Title = "( " & pSpeGroupNm & ")" & "  미생물 업무나열서"


    '        "미생물업무나열서"
    ' ----------------------
    ' ----------------------

    Call prtTitle(Title, lngCm / 4)

    iposSEQ = lngCm / 2                       '284
    iposWorkNo = iposSEQ + lngCm / 1.2        '757
    iposPtName = iposWorkNo + 3 * lngCm       '1891
    iposPtID = iposPtName + lngCm + lngCm / 1.3 + 80 '환자성명이 길경우 겹쳐서 80만 더했음 Modify By Legends 2003/08/11 졸라게 더운 여름날에
    iposSAge = iposPtID + 1.5 * lngCm 'iposPtID + 2 * lngCm
    iposIO = iposSAge + lngCm                 '4311
    iposSF = iposIO + 1.5 * lngCm 'iposIO + 0.9 * lngCm             '4765
    iposTestCD = iposSF + 0.8 * lngCm
    iposRcv = iposTestCD + 4 * lngCm + 80 '검사항목을 좀 넓게 보이게 할려고 80만더했음 Modify By Legends 2003/08/11 졸라게 더운 여름날에
    iposSpccd = iposRcv + 2.2 * lngCm
    'iposAnti = iposSpccd + 1 * lngcm

    iCurY = iCurY + lngCm / 2


    Call DrawLine(lngCm / 2, iCurY, iPageWidth - lngCm / 2, iCurY, "solid", 2)
    Call ChangeLine(iCurY, lngCm / 4)
    Call WriteStr(iCurY, iposSEQ, "SEQ", iCurY, 0)
    Call WriteStr(iCurY, iposWorkNo, "Lab  No", iCurY, 0)
    Call WriteStr(iCurY, iposPtName, "환자성명", iCurY, 0)
    Call WriteStr(iCurY, iposPtID, "환자ID", iCurY, 0)
    Call WriteStr(iCurY, iposSAge, "S/Age", iCurY, 0)
    Call WriteStr(iCurY, iposIO, "I/O", iCurY, 0)
    Call WriteStr(iCurY, iposSF, "S/F", iCurY, 0)
    Call WriteStr(iCurY, iposTestCD, "검사항목", iCurY, 0)
    Call WriteStr(iCurY, iposRcv, "검체도착시간", iCurY, 0)
    Call WriteStr(iCurY, iposSpccd, "검체", iCurY, lngCm / 2)
'    Call WriteStr(iCurY, iposAnti, "항생제", iCurY, lngcm / 2)
    Call DrawLine(lngCm / 2, iCurY, iPageWidth - lngCm / 2, iCurY, "solid", 2)
    Call ChangeLine(iCurY, lngCm / 4)
    'iCurY = , 10268    '2090 , 9650
End Sub
Public Sub prtPageNum()

    Dim oldX As Integer, oldY As Integer
    Dim sDate As String, sTime As String
    Dim lngPosY As Long

    sDate = Format(Now, "YYYY/MM/DD")
    sTime = Format(Now, "HH:MM:SS")
    oldX = Printer.CurrentX
    oldY = Printer.CurrentY

    If pWorksheet2 And (iPageCnt Mod 2) = 0 Then
        lngPosY = iPageHeight / 2
    Else
        lngPosY = 0
    End If
    
    Printer.CurrentX = iPageWidth - 4 * lngCm
    Printer.CurrentY = lngPosY
    Printer.Print "P A G E  : " & iPageCnt

    Printer.CurrentX = iPageWidth - 4 * lngCm
    Printer.CurrentY = lngPosY + Printer.TextHeight("P A G E") + lngCm / 6
    Printer.Print "RUN-DATE : " & sDate

    Printer.CurrentX = iPageWidth - 4 * lngCm
    Printer.CurrentY = lngPosY + Printer.TextHeight("P A G E") + lngCm / 6 + _
                                 Printer.TextHeight("RUN-DATE") + lngCm / 6
    Printer.Print "RUN-TIME : " & sTime

    Printer.CurrentX = oldX
    Printer.CurrentY = oldY

End Sub

Public Sub prtTerm()
    Dim oldX As Integer, oldY As Integer
    Dim sSqlGetMedia As String
    Dim sSqlGetWorkDt As String
    Dim rsMediaCD As Recordset
    Dim rsWorkDt As Recordset
    Dim iXpos As Integer
    Dim strMedia As String
    Dim i%
    Dim lngPosY As Long

    If pWorksheet2 And (iPageCnt Mod 2) = 0 Then
        lngPosY = iPageHeight / 2
    Else
        lngPosY = 0
    End If
    
    sSqlGetWorkDt = " select workdt , worktm" & _
                    " from " & T_LAB401 & _
                    " where " & DBW("wscd=", Trim(pSpeGroupCd)) & _
                    " and " & DBW("wsunit=", Trim(pWsUnit))
                    
    Set rsWorkDt = New Recordset
    rsWorkDt.Open sSqlGetWorkDt, DBConn

    If rsWorkDt.EOF = True Then
        Set rsWorkDt = Nothing
        Exit Sub
    End If

    oldX = Printer.CurrentX
    oldY = Printer.CurrentY

    Printer.CurrentX = lngCm
    Printer.CurrentY = lngPosY + (1.3 * lngCm)
    Printer.Print "작업일시    : " & Trim("" & rsWorkDt.Fields("workdt").Value) & _
                  Chr(vbKeyTab) & CvtTmFormat(Trim("" & rsWorkDt.Fields("worktm").Value))

'    Printer.CurrentX = lngcm
'    Printer.CurrentY = 1.3 * lngcm + Printer.TextHeight("작업일 : ")
'    Printer.Print "WorkSheet Unit    : " & Trim(pWsUnit)

    sSqlGetMedia = " select cdval2 as MediaCD" & _
                   " from  " & T_LAB031 & _
                   " where " & DBW("cdindex=", LC2_SpcMedia) & _
                   " and   " & DBW("cdval1=", Trim(pSpeGroupCd))
                    
    Set rsMediaCD = New Recordset
    rsMediaCD.Open sSqlGetMedia, DBConn
    
    If rsMediaCD.EOF = True Then
        Set rsMediaCD = Nothing
        Exit Sub
    End If

    Printer.CurrentX = lngCm
    Printer.CurrentY = lngPosY + 1.5 * lngCm + Printer.TextHeight("worksheet unit : ")
    Printer.Print "배지코드   :"

    strMedia = ""
    If Not rsMediaCD.EOF Then
        For i = 1 To rsMediaCD.RecordCount
            strMedia = strMedia & "" & rsMediaCD.Fields("MediaCD").Value & ","
            rsMediaCD.MoveNext
        Next i
    End If
    Set rsMediaCD = Nothing
    
    If Len(strMedia) > 0 Then strMedia = Mid(strMedia, 1, Len(strMedia) - 1)
    
    Printer.CurrentX = 3 * lngCm + iXpos
    Printer.CurrentY = lngPosY + 1.5 * lngCm + Printer.TextHeight("worksheet unit : ")
    Printer.Print strMedia

    Printer.CurrentX = oldX
    Printer.CurrentY = oldY
End Sub

Private Sub DspSeq(SeqNum As Integer)
    Call WriteStr(iCurY, iposSEQ, CStr(SeqNum), iCurY, 0)
End Sub

Private Sub DspWorkno(sWorkarea As String, sAccDt As String, iAccSeq As Long)
    Dim sWorkno As String
    Dim ioldfontsize As Integer

    sWorkno = sWorkarea & " - " & DelFirst2Chr(sAccDt) & " -"
    Call WriteStr(iCurY, iposWorkNo, sWorkno, iCurY, 0)

    ioldfontsize = Printer.FontSize
    Printer.FontBold = True
    Call WriteStr(iCurY, iposWorkNo + Printer.TextWidth(sWorkno), _
                  Trim(CStr(iAccSeq)), iCurY, 0)
    Printer.FontBold = False

    Printer.FontSize = ioldfontsize
End Sub

Private Sub DspLab201Data(LabNo As tLabno)
    Dim rslab201Data    As Recordset
    Dim sSqlGetlab201   As String
    Dim S_Age           As String
    Dim strLocation     As String   '진료과/병동
    
    sSqlGetlab201 = " select a.ptid, a.ageday, a.deptcd, a.wardid, a.hosilid, a.storecd, " & _
                    "   a.rcvdt, a.rcvtm , a.spccd, b.field3 as spcnm, c." & F_SEX & " as sex" & _
                    ", a.spcyy || trim( to_char(a.spcno, '000000000') ) as barcode " & _
                    " from  " & T_HIS001 & " c," & T_LAB201 & " a, " & T_LAB032 & " b " & _
                    " where " & DBW("a.workarea=", LabNo.sWorkarea) & _
                    " and   " & DBW("a.accdt   =", LabNo.sAccDt) & _
                    " and   " & DBW("a.accseq  =", LabNo.iAccSeq) & _
                    " and   " & DBW("b.cdindex =", LC3_Specimen) & _
                    " and   b.cdval1 = a.spccd " & _
                    " and   a.ptid = c." & F_PTID
                    
    Set rslab201Data = New Recordset
    rslab201Data.Open sSqlGetlab201, DBConn

    Call WriteStr(iCurY, iposPtID, "" & rslab201Data.Fields("ptid").Value, iCurY, 0)
    S_Age = rslab201Data.Fields("sex").Value & ""
    If IsNumeric(S_Age) Then
        S_Age = Choose((Val(S_Age) Mod 2) + 1, "F", "M")
    End If
    S_Age = S_Age & "/" & ("" & rslab201Data.Fields("ageday").Value \ 365 + 1)
    
    '## 5.0.4: 이상대(2004-12-28)
    '   - 병동환자인 경우 병동,병실을 출력
    If rslab201Data.Fields("wardid").Value & "" = "" Then
        strLocation = rslab201Data.Fields("deptcd").Value & ""
    Else
        strLocation = rslab201Data.Fields("wardid").Value & "" & "-" & rslab201Data.Fields("hosilid").Value & ""
    End If
    
    Call WriteStr(iCurY, iposSAge, S_Age, iCurY, 0)
    Call WriteStr(iCurY, iposIO, strLocation, iCurY, 0)
    Call WriteStr(iCurY, iposSF, "" & rslab201Data.Fields("storecd").Value, iCurY, 0)
    Call WriteStr(iCurY, iposRcv, DelFirst2Chr("" & rslab201Data.Fields("rcvdt").Value), iCurY, 0)
    Call WriteStr(iCurY, iposRcv + lngCm + lngCm / 6, CvtTmFormat("" & rslab201Data.Fields("rcvtm").Value), iCurY, 0)
    Call WriteStr(iCurY, iposSpccd, "" & rslab201Data.Fields("spcnm").Value, iCurY, 0)
'    Call WriteStr(iCurY + 1, iposRcv, rslab201Data.Fields("barcode").Value, iCurY + 1, 0)

    Set rslab201Data = Nothing
End Sub
Private Function DelFirst2Chr(sStr As String) As String
    DelFirst2Chr = Trim(Mid(sStr, 3, Len(sStr) - 2))
End Function
Private Function DelLast2Chr(sStr As String) As String
    DelLast2Chr = Trim(Mid(sStr, 1, Len(sStr) - 2))
End Function
Private Function CvtTmFormat(sStr As String) As String
    Dim Time As String
    Dim Hour As String
    Dim Min As String

    Time = DelLast2Chr(sStr)
    Hour = DelLast2Chr(Time)
    Min = DelFirst2Chr(Time)

    CvtTmFormat = Hour & ":" & Min

End Function

Private Sub DspPtnm(LabNo As tLabno)
    Dim sSqlGetPtnm As String
    Dim rsGetPtnm As Recordset


    sSqlGetPtnm = " select " & F_PTNM & " as ptnm " & _
                    " from " & T_HIS001 & _
                    " where " & F_PTID & " = " & _
                    "       (select ptid " & _
                    "        from " & T_LAB201 & _
                    "        where " & DBW("workarea=", LabNo.sWorkarea) & _
                    "           and " & DBW("accdt=", LabNo.sAccDt) & _
                    "           and " & DBW("accseq=", LabNo.iAccSeq) & ")"

    Set rsGetPtnm = New Recordset
    rsGetPtnm.Open sSqlGetPtnm, DBConn
    
    If rsGetPtnm.EOF = True Then Exit Sub

    Call WriteStr(iCurY, iposPtName, "" & rsGetPtnm.Fields("ptnm").Value, iCurY, 0)

    Set rsGetPtnm = Nothing

End Sub

Private Sub DspTestCD(LabNo As tLabno)
    Dim sSqlGetTestCD   As String
    Dim rsGetTestCD     As Recordset

    Dim sSqlGetTestNm   As String
    Dim rsGetTestNm     As Recordset
    Dim i%, oldipostestcd%

    Dim LastRst         As tLastRst
    Dim strLastRst      As String

    Dim sICSString      As String
    
    sICSString = ICSLabNoString(LabNo.sWorkarea, LabNo.sAccDt, LabNo.iAccSeq, enICSNum.LIS_ALL)
    
    '## 이전쿼리
'    sSqlGetTestCD = " select b.ptid,a.testcd,b.spccd, a.rsttype " & _
'                    " from " & T_LAB404 & " a, " & T_LAB201 & " b " & _
'                    " where " & DBW("a.workarea=", LabNo.sWorkarea) & _
'                    " and " & DBW("a.accdt=", LabNo.sAccDt) & _
'                    " and " & DBW("a.accseq=", LabNo.iAccSeq) & _
'                    " and ( a.detailfg is null or a.detailfg = '' )" & _
'                    " and " & DBW("a.stscd<", enStsCd.StsCd_LIS_FinRst) & _
'                    " and a.workarea=b.workarea " & _
'                    " and a.accdt=b.accdt " & _
'                    " and a.accseq=b.accseq " & _
'                    " union " & _
'                    " select b.ptid,a.testcd,b.spccd, a.rsttype " & _
'                    " from " & T_LAB404 & " a, " & T_LAB201 & " b " & _
'                    " where " & DBW("a.workarea=", LabNo.sWorkarea) & _
'                    " and " & DBW("a.accdt=", LabNo.sAccDt) & _
'                    " and " & DBW("a.accseq=", LabNo.iAccSeq) & _
'                    " and   a.detailfg <> ' ' and  a.detailfg is not null " & _
'                    " and " & DBW("a.rstdiv=", "*") & _
'                    " and " & DBW("a.stscd<", enStsCd.StsCd_LIS_FinRst) & _
'                    " and a.workarea=b.workarea " & _
'                    " and a.accdt=b.accdt " & _
'                    " and a.accseq=b.accseq "

    '## Modify 이상대(2004-12-04)
    '   - 해당 Worksheet에 해당하는 검사명을 조회하기 위해 결과유형을 조건에 추가
    sSqlGetTestCD = " select b.ptid,a.testcd,b.spccd, a.rsttype " & _
                    " from " & T_LAB404 & " a, " & T_LAB201 & " b " & _
                    " where " & DBW("a.workarea=", LabNo.sWorkarea) & _
                    " and " & DBW("a.accdt=", LabNo.sAccDt) & _
                    " and " & DBW("a.accseq=", LabNo.iAccSeq) & _
                    " and a.rsttype in (" & GetRTypes(pSpeGroupField4) & ")" & _
                    " and ( a.detailfg is null or a.detailfg = '' )" & _
                    " and " & DBW("a.stscd<=", enStsCd.StsCd_LIS_FinRst) & _
                    " and a.workarea=b.workarea " & _
                    " and a.accdt=b.accdt " & _
                    " and a.accseq=b.accseq " & _
                    " union " & _
                    " select b.ptid,a.testcd,b.spccd, a.rsttype " & _
                    " from " & T_LAB404 & " a, " & T_LAB201 & " b " & _
                    " where " & DBW("a.workarea=", LabNo.sWorkarea) & _
                    " and " & DBW("a.accdt=", LabNo.sAccDt) & _
                    " and " & DBW("a.accseq=", LabNo.iAccSeq) & _
                    " and a.rsttype in (" & GetRTypes(pSpeGroupField4) & ")" & _
                    " and   a.detailfg <> ' ' and  a.detailfg is not null " & _
                    " and " & DBW("a.rstdiv=", "*") & _
                    " and " & DBW("a.stscd<=", enStsCd.StsCd_LIS_FinRst) & _
                    " and a.workarea=b.workarea " & _
                    " and a.accdt=b.accdt " & _
                    " and a.accseq=b.accseq "

    Set rsGetTestCD = New Recordset
    rsGetTestCD.Open sSqlGetTestCD, DBConn
    
    If rsGetTestCD.EOF = True Then
        GoTo EndLine
        Exit Sub
    End If

    oldipostestcd = iposTestCD  '**********************************************************

    For i = 1 To rsGetTestCD.RecordCount
        LastRst.sPtid = rsGetTestCD.Fields("ptid").Value & ""
        LastRst.sTestCd = rsGetTestCD.Fields("testcd").Value & ""
        LastRst.sSpcCd = rsGetTestCD.Fields("spccd").Value & ""

        strLastRst = GetNoGrowthLatestRst(LastRst)

        sSqlGetTestNm = " select abbrnm5 " & _
                        " from " & T_LAB001 & _
                        " where " & DBW("testcd=", "" & rsGetTestCD.Fields("testcd").Value)
        Set rsGetTestNm = Nothing
        Set rsGetTestNm = New Recordset
        rsGetTestNm.Open sSqlGetTestNm, DBConn
        
        If Trim(strLastRst) <> "" Then
            Call WriteStr(iCurY, iposTestCD, "" & rsGetTestNm.Fields("abbrnm5").Value & _
                                             "(" & strLastRst & ")", iCurY, 0)
        Else
            Call WriteStr(iCurY, iposTestCD, "" & rsGetTestNm.Fields("abbrnm5").Value, iCurY, 0)
        End If

'        iposTestCD = iposTestCD + 2 * lngCm
        iposTestCD = iposTestCD + 2 * 400

        If pSpeGroupField4 = "GC" Then             ' newPage일 경우
            If pSpeGroupCd = "BL" Or Not pWorksheet2 Then   '2000.08.08 수정
                If i Mod 9 = 0 Then
                    Printer.NewPage
                    Call PrtHeader
                    Call prtPageNum
                    Call prtTerm
                End If
            Else
'                If i Mod 5 = 0 Then
                If i Mod 2 = 0 Then
                    Printer.NewPage
                    Call PrtHeader
                    Call prtPageNum
                    Call prtTerm
                End If
            End If
        ElseIf pSpeGroupField4 <> "GC" And i Mod 15 = 0 Then
            Printer.NewPage
            Call PrtHeader
            Call prtPageNum
            Call prtTerm
        End If

        If pSpeGroupCd <> "BL" And pWorksheet2 Then
            If pSpeGroupField4 = "GC" And "" & rsGetTestCD.Fields("rsttype").Value = "G" Then
                Call Dsp_GC_TestNm("" & rsGetTestNm.Fields("abbrnm5").Value, iCurY)
            End If
        End If

        rsGetTestCD.MoveNext
    Next
    
    
    '처방 리마크 찍어주기
    Dim Rs          As Recordset
    Dim arytmp()    As String
    Dim strTmp      As String
    Dim sSQL        As String
    
    sSQL = " select mesg from " & T_LAB102 & _
           " where" & _
                     DBW("workarea=", LabNo.sWorkarea) & _
           " and " & DBW("accdt=", LabNo.sAccDt) & _
           " and " & DBW("accseq=", LabNo.iAccSeq) & _
           " and (mesg<>' ' or mesg is not null)"
           
    Set Rs = New Recordset
    Rs.Open sSQL, DBConn
    
    If Not Rs.EOF Then
        strTmp = Trim(Rs.Fields("mesg").Value & "")
        If strTmp <> "" Then
            arytmp() = Split(strTmp, vbCrLf)
            For i = LBound(arytmp) To UBound(arytmp)
                If arytmp(i) = "" Then Exit For
                
                iCurY = iCurY + lngCm / 2
                If i = 0 Then
                    Call WriteStr(iCurY, iposWorkNo, "처방 Remark:" & Space(1) & arytmp(i), iCurY, 1)
                Else
                    Call WriteStr(iCurY, iposWorkNo, arytmp(i), iCurY, 1)
                End If
            Next
        End If
    End If
    Set Rs = Nothing
    
    iposTestCD = oldipostestcd  '**********************************************************
    
EndLine:
    If sICSString <> "" Then
        iCurY = iCurY + lngCm / 4
    End If
    
    If (pSpeGroupField4 = "GC") And (pSpeGroupCd <> "BL") And pWorksheet2 Then
        iCurY = iCurY + 8.5 * lngCm
        Call DrawLine(lngCm / 2, iCurY, iPageWidth - lngCm / 2, iCurY, "dot", 1)
        Call ChangeLine(iCurY, lngCm / 4)
    Else
        '업무나열서 사이공간을 넓히려면+ lngCm*??
        iCurY = iCurY + lngCm
        Call DrawLine(lngCm / 2, iCurY, iPageWidth - lngCm / 2, iCurY, "dot", 1)
        Call ChangeLine(iCurY, lngCm / 4)
    End If

    Set rsGetTestCD = Nothing
    Set rsGetTestNm = Nothing
End Sub

Private Function GetAccNoOfLatestRst(LastRst As tLastRst) As String
    Dim Rs As Recordset
    Dim strSQL As String
   
    strSQL = " select a.workarea, a.accdt, a.accseq, a.vfydt, a.vfytm " & _
             " from " & T_LAB404 & " a, " & T_LAB201 & " b " & _
             " where " & DBW("a.ptid = ", LastRst.sPtid) & _
             " and  (a.vfydt<>'' or a.vfydt is not null ) " & _
             " and " & DBW("a.testcd=", LastRst.sTestCd) & _
             " and   " & DBW("a.stscd >= ", enStsCd.StsCd_LIS_FinRst) & _
             " and     b.workarea = a.workarea " & _
             " and     b.accdt = a.accdt " & _
             " and     b.accseq = a.accseq " & _
             " and   " & DBW("b.spccd = ", LastRst.sSpcCd) & _
             " order by b.rcvdt desc, b.rcvtm desc "
    
    Set Rs = New Recordset
    Rs.Open strSQL, DBConn
    
    If Rs.EOF Then
        GetAccNoOfLatestRst = ""
    Else
        GetAccNoOfLatestRst = Rs.Fields("workarea").Value & "" & "-" & _
                              Rs.Fields("accdt").Value & "" & "-" & _
                              Rs.Fields("accseq").Value & "" & "-" & _
                              Rs.Fields("vfydt").Value & "" & "-" & _
                              Rs.Fields("vfytm").Value & ""
    End If
    
    Set Rs = Nothing
End Function

Private Function GetNoGrowthLatestRst(LastRst As tLastRst) As String
    Dim Rs As Recordset
    Dim strSQL As String
    Dim strAccNo As String
    Dim strWA As String
    Dim strAccdt As String
    Dim strAccseq As String
    
    GetNoGrowthLatestRst = ""
    
    strAccNo = GetAccNoOfLatestRst(LastRst)
    
    If strAccNo <> "" Then
        strWA = medGetP(strAccNo, 1, "-")
        strAccdt = medGetP(strAccNo, 2, "-")
        strAccseq = medGetP(strAccNo, 3, "-")
        
        strSQL = " select rstcd from " & T_LAB404 & _
                 " where " & DBW("workarea=", strWA) & _
                 " and " & DBW("accdt=", strAccdt) & _
                 " and " & DBW("accseq=", strAccseq) & _
                 " and " & DBW("testcd=", LastRst.sTestCd)
        
        Set Rs = New Recordset
        Rs.Open strSQL, DBConn
        
        If Rs.EOF = False Then
            GetNoGrowthLatestRst = Rs.Fields("rstcd").Value & ""
        End If
        
        Set Rs = Nothing
    End If
End Function

Private Sub Dsp_GC_TestNm(TestNm As String, ByVal icuryTmp)

     'Call WriteStr(icuryTmp + 2.5 * lngcm, lngcm / 2, TestNm & " : ", icuryTmp, 0)

    If pSpeGroupCd = "BL" Or Not pWorksheet2 Then
        Printer.CurrentY = icuryTmp + 0.5 * lngCm
    Else
        Printer.CurrentY = icuryTmp + 3.8 * lngCm
    End If
    Printer.CurrentX = iposWorkNo
    'iNextY = Printer.CurrentY + iSpace
    Printer.Print TestNm & " : "

End Sub

Private Sub PrtBody()
    Dim LabNo As tLabno
    Dim sSqlGetLabno As String
    Dim rsLabno As Recordset
    Dim i%
    Dim objSql As New clsLISSqlStatement
    Dim strTemp As String
    Dim strBarCode As String

    '2001/11/21 아래문장 교체
    'FUNC_CONVERT("int", "d.text1") & " RptSeq" ==> FUNC_CONVERT("char(255)", "d.text1") & " RptSeq"
    If ObjSysInfo.dbtype = 0 Then
        strTemp = FUNC_CONVERT("num", "d.text1") & " RptSeq"
    Else
        strTemp = FUNC_CONVERT("char", "d.text1") & " RptSeq"
    End If
    sSqlGetLabno = " select distinct a.workarea, a.accdt, a.accseq, c.spccd, c.rcvdt, c.rcvtm, " & strTemp & _
                    " from " & T_LAB032 & " d, " & T_LAB201 & " c, " & T_LAB404 & " b, " & T_LAB402 & " a " & _
                    " where " & DBW("a.wscd=", Trim(pSpeGroupCd)) & _
                    " and   " & DBW("a.wsunit=", Trim(pWsUnit)) & _
                    " and   b.workarea = a.workarea " & _
                    " and   b.accdt = a.accdt " & _
                    " and   b.accseq = a.accseq " & _
                    " and   c.workarea = b.workarea " & _
                    " and   c.accdt  = b.accdt " & _
                    " and   c.accseq = b.accseq " & _
                    " and " & DBW("d.cdindex = ", LC3_Specimen) & _
                    " and   d.cdval1 = c.spccd"
    '2000.08.08 추가 : Nogrowth Batch등록에서 보류리스트의 Worksheet을 출력할 경우...
    If pWorksheet2 Then sSqlGetLabno = sSqlGetLabno & "   AND a.genfg='" & MWS_Holding & "' AND a.selfg <> '" & MWS_Final & "' "
    If pStainWorksheet Then sSqlGetLabno = sSqlGetLabno & "   AND a.genfg='" & MWS_Holding & "' AND a.selfg <> '" & MWS_Final & "' "

    sSqlGetLabno = sSqlGetLabno & " order by workarea, accdt, accseq "
        
    Set rsLabno = New Recordset
    rsLabno.Open sSqlGetLabno, DBConn
    
    For i = 1 To rsLabno.RecordCount

        LabNo.sWorkarea = Trim("" & rsLabno.Fields("workarea").Value)
        LabNo.sAccDt = Trim("" & rsLabno.Fields("accdt").Value)
        LabNo.iAccSeq = Trim("" & rsLabno.Fields("accseq").Value)

        If iCurY > iPageHeight - 2 * lngCm Then  ' newPage일 경우
            Printer.NewPage
            If Not pWorksheet2 Then
                Call PrtHeader
                Call prtPageNum
                Call prtTerm
                Call Print_WaterMark
            End If
        End If

        If pWorksheet2 Then
            Call PrtHeader
            Call prtPageNum
            Call prtTerm
        End If
        Dim sICSString As String
        
        Call DspSeq(i)
        sICSString = ICSLabNoString(LabNo.sWorkarea, LabNo.sAccDt, LabNo.iAccSeq, enICSNum.LIS_ALL)
        If sICSString <> "" Then
            Call WriteStr(iCurY, iposWorkNo, "Infection : " & sICSString, iCurY, 0)
            Call ChangeLine(iCurY, lngCm / 4)
        End If
        
        Call DspWorkno(LabNo.sWorkarea, LabNo.sAccDt, LabNo.iAccSeq)    'WORKAREA
        Call DspLab201Data(LabNo)                                       '환자 검체정보
        Call DspPtnm(LabNo)                                             '환자명
        Call DspTestCD(LabNo)                                           '
'        Call DspAntiBiotic
 '       Call ChangeLine(lngcm * 4.5)
   '     Call ChangeLine(icury, lngcm * 4.5)
        rsLabno.MoveNext
    Next i

    Set rsLabno = Nothing
    Set objSql = Nothing
End Sub

Public Function GetRTypes(ByVal pSG As String) As String
    Dim dsRT        As New Recordset
    Dim sSQL        As String
    Dim sRType()    As String

    
    sSQL = " SELECT field2 FROM " & T_LAB032 & _
           " WHERE " & DBW("cdindex = ", LC3_MWSKinds) & _
           " AND   " & DBW("cdval1  = ", pSG)
                    
    GetRTypes = ""
    dsRT.Open sSQL, DBConn
    
    If dsRT.EOF Then GoTo Nodata
     
    Dim i As Integer
    
    sRType = Split(dsRT.Fields("field2").Value & "", ",")
    For i = LBound(sRType) To UBound(sRType)
        sRType(i) = "'" & sRType(i) & "'"
    Next
    GetRTypes = Join(sRType, ",")
     
Nodata:
    Set dsRT = Nothing
End Function
