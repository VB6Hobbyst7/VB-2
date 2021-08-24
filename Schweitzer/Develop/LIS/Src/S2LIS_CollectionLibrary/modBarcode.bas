Attribute VB_Name = "modBarcode"
   
'-----------------------------------------------------------------------------------
'  Command ds : Syntax - '\1Bds' & p1 & p2 & p3 & p4 & p5 & p6 & p7 & p8 & p9 & p10 & p11
'                       p1 - Format No. (2 byte)
'                       p2 - Element No. (2 byte)
'                       p3 - ST (2 byte) : 01=변경안함, 02=변경
'                       p4 - ST-X (4 byte) : Horizontal start position (X) in dots. (0000~0447)
'                       p5 - ST-Y (4 byte) : Vertical start position (Y) in dots. (0000~2200)
'                       p6 - Length (2 byte) : 00~28
'                       p7 - Font Size (2 byte) : Mg-X(1~6), Mg-Y(1~6)
'                       p8 - Rotate (2 byte) : 16 가지
'                       p9 - Reverse (1 byte) : 0 or 1
'                       p10 - 한글Font (1 byte): 0=바탕,1=굴림
'                       p11 - Bold (1 byte) : 0 or 1
'-----------------------------------------------------------------------------------
'  Command bs : Syntax - '\1Bbs' & p1 & p2 & p3 & p4 & p5 & p6 & p7 & p8 & p9 & p10 & p11 & p12
'                       p1 - Format No. (2 byte)
'                       p2 - Element No. (2 byte)
'                       p3 - ST (2 byte) : 01=변경안함, 02=변경
'                       p4 - ST-X (4 byte) : Horizontal start position (X) in dots. (0000~0447)
'                       p5 - ST-Y (4 byte) : Vertical start position (Y) in dots. (0000~2200)
'                       p6 - Length (2 byte) : 00~28
'                       p7 - Height (4 byte) : Bar code height in dots.
'                       p8 - Symbology (2 byte) : 00=Code39, 01=Checkcode39, 02=Intereaved2of5, 03=Matrix2of5.
'                                                            04=Industrial2of5, 05=Codabar, 06=NW-7hex, 07=Upc-A, ..~16
'                       p9 - N.Thick (1 byte) : Narrow bar width in dots (0~7)
'                       p10 - N.W Ratio (1 byte) : Wide bar ratio. (0~3:2,2.5,3,3.5)
'                       p11 - Rotation (1 byte)
'                       p12 - Print human readable code (2 byte) : 00 ~ ?
'-----------------------------------------------------------------------------------
'  Command ls : Syntax - '\1Bls' & p1 & p2 & p3 & p4 & p5 & p6 & p7
'                       p1 - Format No. (2 byte)
'                       p2 - Element No. (2 byte)
'                       p3 - ST-X (4 byte) : Horizontal start position (X) in dots. (0000~0447)
'                       p4 - ST-Y (4 byte) : Vertical start position (Y) in dots. (0000~2200)
'                       p5 - Horizontal length in dots (4 byte).
'                       p6 - Vertical length in dots (4 byte).
'                       p7 - Thick (4 byte) : 0000 ~ 0007
'-----------------------------------------------------------------------------------
'  Command q :  Syntax - '\1q" & p1
'                       p1 - Quantity (4 byte) : Number of copies of each label. (0001~9999)
'-----------------------------------------------------------------------------------

Private Const FormatNo = "09"

Private Const LabelWidth = "0320"
Private Const LabelLength = "0184"
Private Const LabelTotLength = "0208"
Private Const GapLength = 24
Private Const PosYbar = "0035"
Private Const PosY1 = "0001"
Private Const PosY2 = "0019"
Private Const PosY3 = "0055"
Private Const PosY4 = "0120"
Private Const PosY5 = "0125"
Private Const PosY6 = "0140"
Private Const PosY7 = "0150"
Private Const PosY8 = "0165"
Private Const PosY9 = "0035"
Private Const PosYA = "0082"
Private Const PosYB = "0100"

Private Const PosXbld = "0150"       'building
Private Const PosXbar = "0155"       'barcode
Private Const PosXwa = "0190"        'workarea
Private Const PosXdt = "0260"          'coldt
Private Const PosXseq = "0370"       'accseq
Private Const PosXsno = "0240"       'spc no
Private Const PosXward = "0400"       'ward
Private Const PosXstore = "0420"     'storecd
Private Const PosXpnm = "0150"      'ptnm
Private Const PosXpid = "0235"        'ptid
Private Const PosXspc = "0340"       'spcnm
Private Const PosXtest = "0150"       'testnm
Private Const PosXstat = "0150"       'stat

Private Const StFg = "00"
Private Const FontDF = "0"
Private Const FontSS = "1"
Private Const FontMd = "2"
Private Const FontLg = "3"
Private Const FontKor1 = "0"  '바탕체
Private Const FontKor2 = "1"  '굴림체
Private Const NoRot = "00"
Private Const Rot90 = "01"
Private Const Rot180 = "02"
Private Const Rot270 = "03"
Private Const Reverse = "1"
Private Const Normal = "0"
Private Const Code39 = "00"
Private Const Code39C = "01"
Private Const Code2of5 = "02"
Private Const BarHeight = "0080"
Private Const Readable = "01"
Private Const NotReadable = "00"
Private Const NarrowBar = "1"
Private Const NWRatio = "1"
Private Const BarLength = "12"
Private Const Bold = "1"
   
Private mvarWorkArea As String 'work area
Private mvarColDt As String '채혈일
Private mvarAccSeq As String 'accession sequence
Private mvarStatFg As String '응급여부
Private mvarSpcNo As String '검체번호
Private mvarPtId As String '환자ID
Private mvarPtNm As String '환자명
Private mvarSpcNm As String '검체명
Private mvarStoreCd As String '보관구분
Private mvarWardId As String '병동ID
Private mvarLocation As String '검체전달location
Private mvarTestNames As String '검사명
Private mvarCopyCount As Integer '출력장수


Public Sub Clear()
   mvarWorkArea = ""
   mvarColDt = ""
   mvarAccSeq = ""
   mvarStatFg = ""
   mvarSpcYY = ""
   mvarSpcSeq = ""
   mvarPtId = ""
   mvarPtNm = ""
   mvarSpcNm = ""
   mvarStoreCd = ""
   mvarWardId = ""
   mvarLocation = ""
   mvarTestNames = ""
   mvarCopyCount = 0
End Sub


Public Sub Label_PortOpen()
    With frmControls
        If .MyComm.PortOpen Then Exit Sub
        .MyComm.CommPort = 2
        .MyComm.Settings = "9600,N,8,1"
        .MyComm.InputLen = 8192
        
        If Not .MyComm.PortOpen Then .MyComm.PortOpen = True
    End With
End Sub

        

Public Sub Label_PortClose()
        
    If frmControls.MyComm.PortOpen Then frmControls.MyComm.PortOpen = False

End Sub

Public Sub GetBarInfo(ByVal strOrdDiv As String)

    '바코드 출력양식 읽어오기
    Select Case strOrdDiv
    Case "A"
        If Not blnAPSBarFg Then
            Set objAPSbarcode.MyDb = dbConn
            objAPSbarcode.ProjectCd = "APS"
            Call objAPSbarcode.GetBarConfig
            blnAPSBarFg = True
        End If
    Case "B"
        If Not blnBBSBarFg Then
            Set objBBSbarcode.MyDb = dbConn
            objBBSbarcode.ProjectCd = "BBS"
            Call objBBSbarcode.GetBarConfig
            blnBBSBarFg = True
        End If
    Case "L"
        If Not blnLISBarFg Then
            Set objLISbarcode.MyDb = dbConn
            objLISbarcode.ProjectCd = "LIS"
            Call objLISbarcode.GetBarConfig
            blnLISBarFg = True
        End If
    End Select

End Sub


Public Sub Label_PrintOut(ByVal strOrdDiv As String, ByVal Location As Variant, ByVal WorkArea As Variant, _
                          ByVal ColDt As Variant, ByVal AccSeq As Variant, ByVal SpcNo As Variant, _
                          ByVal PtId As Variant, ByVal PtNm As Variant, ByVal SpcNm As Variant, _
                          ByVal StoreCd As Variant, ByVal StatFg As Variant, ByVal WardId As Variant, _
                          ByVal OrdDt As Variant, ByVal ColTm As Variant, ByVal TestNames As Variant, _
                          ByVal CopyCount As Variant, _
                          Optional ByVal AccFg As Boolean = False, Optional ByVal FzFg As Boolean = False)
    Dim barString As String
    Dim FileNo As Long
    'Dim MyComm As Object
    Dim PkSize As Integer

    On Error GoTo Skip
   
    Call GetBarInfo(strOrdDiv)
   
    Select Case strOrdDiv
    Case "A":
        Call objAPSbarcode.Label_PrintOut(Location, WorkArea, ColDt, AccSeq, SpcNo, PtId, PtNm, _
                        SpcNm, StoreCd, StatFg, WardId, OrdDt, ColTm, TestNames, CopyCount, AccFg, FzFg)
    Case "B":
        Call objBBSbarcode.Label_PrintOut(Location, WorkArea, ColDt, AccSeq, SpcNo, PtId, PtNm, _
                        SpcNm, StoreCd, StatFg, WardId, OrdDt, ColTm, TestNames, CopyCount, AccFg, FzFg)
    Case "L"
        Call objLISbarcode.Label_PrintOut(Location, WorkArea, ColDt, AccSeq, SpcNo, PtId, PtNm, _
                        SpcNm, StoreCd, StatFg, WardId, OrdDt, ColTm, TestNames, CopyCount, AccFg, FzFg)
    End Select
   
   
   'Set MyComm = frmComm.MSComm1
      
'   barString = Label_String(WorkArea, ColDt, AccSeq, StatFg, SpcNo, PtId, PtNm, SpcNm, StoreCd, WardId, _
'                            Location, TestNames, CopyCount, AccFg, OrdDt, ColTm)
'   PkSize = 250
'
'   With frmControls
'        '.MyComm.PortOpen = True
'        If Not .MyComm.PortOpen Then Call Label_PortOpen
'
'        If Len(barString) > PkSize Then
'            .MyComm.Output = Mid(barString, 1, PkSize)
'            While (Len(barString)) > PkSize
'                   barString = Mid(barString, PkSize + 1)
'                   .MyComm.Output = Mid(barString, 1, PkSize)
'            Wend
'        Else
'            .MyComm.Output = barString
'        End If
'        '.MyComm.PortOpen = False
'   End With
'
Skip:
   'Call Clear
   'Set medMain.MyComm = Nothing

End Sub

Public Function Label_String(ByVal WorkArea As Variant, ByVal AccDt As Variant, ByVal AccSeq As Variant, _
                             ByVal StatFg As Variant, ByVal SpcNo As Variant, _
                             ByVal PtId As Variant, ByVal PtNm As Variant, ByVal SpcNm As Variant, _
                             ByVal StoreCd As Variant, ByVal WardId As Variant, ByVal Location As Variant, _
                             ByVal TestNames As Variant, ByVal CopyCount As Variant, Optional ByVal AccFg As Boolean = False, _
                             Optional ByVal OrdDt As String = "", Optional ByVal ColTm As String = "")

   
   'If IsMissing(WorkArea) Then WorkArea = mvarWorkArea
   'If IsMissing(ColDt) Then ColDt = mvarColDt
   'If IsMissing(AccSeq) Then AccSeq = mvarAccSeq
   'If IsMissing(StatFg) Then StatFg = mvarStatFg
   'If IsMissing(SpcNo) Then SpcNo = mvarSpcNo
   'If IsMissing(PtId) Then PtId = mvarPtId
   'If IsMissing(PtNm) Then PtNm = mvarPtNm
   'If IsMissing(SpcNm) Then SpcNm = mvarSpcNm
   'If IsMissing(StoreCd) Then StoreCd = mvarStoreCd
   'If IsMissing(WardId) Then WardId = mvarWardId
   'If IsMissing(Location) Then Location = mvarLocation
   'If IsMissing(TestNames) Then TestNames = mvarTestNames
   'If IsMissing(CopyCount) Then CopyCount = mvarCopyCount
   
   If CopyCount = 0 Then CopyCount = 1
   If Len(TestNames) > 0 Then TestNames = Mid(TestNames, 1, Len(TestNames) - 1)
   If AccFg Then AccSeq = AccSeq & Space(4 - Len(AccSeq)) & "V"
   SpcNo = AddCheckDigit(CStr(SpcNo))    'check digit 추가
   
   Label_String = ""
   Label_String = Label_String & "\1B@z" & vbCrLf
   Label_String = Label_String & "\1B@f" & FormatNo & vbCrLf
   Label_String = Label_String & "\1Ba" & FormatNo & LabelLength & LabelTotLength & vbCrLf
   Label_String = Label_String & "\1Bf" & FormatNo & vbCrLf
   
   Label_String = Label_String & "\1Bbs" & FormatNo & "02" & StFg & PosXbar & PosYbar & BarLength & BarHeight & Code2of5 & NarrowBar & NWRatio & Normal & NotReadable & vbCrLf    'Barcode Label
   
   If WardId = "ER" And Location = "응급" Then
        Label_String = Label_String & "\1Bds" & FormatNo & "02" & StFg & PosXbld & PosY1 & "04" & FontSS & FontMd & NoRot & Reverse & FontKor2 & Bold & vbCrLf   '건물
   Else
        Label_String = Label_String & "\1Bds" & FormatNo & "02" & StFg & PosXbld & PosY1 & "04" & FontSS & FontMd & NoRot & Normal & FontKor2 & Bold & vbCrLf   '건물
   End If
   Label_String = Label_String & "\1Bds" & FormatNo & "04" & StFg & PosXwa & PosY1 & "02" & FontMd & FontMd & NoRot & Normal & FontKor2 & Normal & vbCrLf   'Workarea
   Label_String = Label_String & "\1Bds" & FormatNo & "06" & StFg & PosXdt & PosY1 & "12" & FontSS & FontSS & NoRot & Normal & FontKor2 & Bold & vbCrLf   'AccDt
   'Label_String = Label_String & "\1Bds" & FormatNo & "06" & StFg & PosXdt & PosY1 & "05" & FontSS & FontSS & NoRot & Normal & FontKor2 & Bold & vbCRLF   '채혈일
   Label_String = Label_String & "\1Bds" & FormatNo & "08" & StFg & PosXseq & PosY1 & "06" & FontMd & FontMd & NoRot & Normal & FontKor2 & Normal & vbCrLf   'AccSeq
   Label_String = Label_String & "\1Bds" & FormatNo & "10" & StFg & PosXsno & PosY2 & "12" & FontSS & FontSS & NoRot & Normal & FontKor2 & Normal & vbCrLf   '검체번호
   'Label_String = Label_String & "\1Bds" & FormatNo & "12" & StFg & PosXstore & PosY3 & "01" & FontMd & FontLg & NoRot & Normal & FontKor2 & Bold & vbCrLf   '보관구분
   Label_String = Label_String & "\1Bds" & FormatNo & "14" & StFg & PosXward & PosYA & "05" & FontSS & FontSS & NoRot & Normal & FontKor2 & Normal & vbCrLf   '처방일
   Label_String = Label_String & "\1Bds" & FormatNo & "16" & StFg & PosXward & PosYB & "05" & FontSS & FontSS & NoRot & Normal & FontKor2 & Normal & vbCrLf   '희망채혈일시
   Label_String = Label_String & "\1Bds" & FormatNo & "18" & StFg & PosXpnm & PosY4 & "00" & FontSS & FontSS & NoRot & Normal & FontKor2 & Normal & vbCrLf   '환자명
   Label_String = Label_String & "\1Bds" & FormatNo & "20" & StFg & PosXpid & PosY5 & "10" & FontSS & FontSS & NoRot & Normal & FontKor2 & Normal & vbCrLf   '환자ID
   Label_String = Label_String & "\1Bds" & FormatNo & "22" & StFg & PosXspc & PosY4 & "10" & FontDF & FontSS & NoRot & Normal & FontKor2 & Bold & vbCrLf   '검체명
   Label_String = Label_String & "\1Bds" & FormatNo & "24" & StFg & PosXtest & PosY6 & "00" & FontSS & FontSS & NoRot & Normal & FontKor2 & Normal & vbCrLf   '검사명
   Label_String = Label_String & "\1Bds" & FormatNo & "26" & StFg & PosXtest & PosY7 & "00" & FontSS & FontSS & NoRot & Normal & FontKor2 & Normal & vbCrLf   '검사명2
   
   Label_String = Label_String & "\1Bds" & FormatNo & "28" & StFg & PosXward & PosY9 & "06" & FontSS & FontSS & NoRot & Normal & FontKor2 & Bold & vbCrLf   'Ward Id
   
   If Trim(StatFg) = "1" Then Label_String = Label_String & "\1Bls" & FormatNo & "02" & PosXstat & PosY8 & "0300" & "0000" & "0007" & vbCrLf   '응급
   
   Label_String = Label_String & "\1Bbw0902" & SpcNo & vbCrLf
   
   Label_String = Label_String & "\1Bdw0902" & Location & vbCrLf
   Label_String = Label_String & "\1Bdw0904" & WorkArea & vbCrLf
   Label_String = Label_String & "\1Bdw0906" & AccDt & vbCrLf
   'Label_String = Label_String & "\1Bdw0906" & ColDt & vbCRLF
   Label_String = Label_String & "\1Bdw0908" & AccSeq & vbCrLf
   Label_String = Label_String & "\1Bdw0910" & SpcNo & vbCrLf
   'Label_String = Label_String & "\1Bdw0912" & StoreCd & vbCrLf
   Label_String = Label_String & "\1Bdw0914" & OrdDt & vbCrLf
   Label_String = Label_String & "\1Bdw0916" & ColTm & vbCrLf
   Label_String = Label_String & "\1Bdw0918" & PtNm & vbCrLf
   Label_String = Label_String & "\1Bdw0920" & PtId & vbCrLf
   Label_String = Label_String & "\1Bdw0922" & SpcNm & vbCrLf

   If Len(TestNames) > 36 Then
      Label_String = Label_String & "\1Bdw0924" & Mid(TestNames, 1, 36) & vbCrLf
      Label_String = Label_String & "\1Bdw0926" & Mid(TestNames, 37) & vbCrLf
   Else
      Label_String = Label_String & "\1Bdw0924" & TestNames & vbCrLf
      Label_String = Label_String & "\1Bdw0926" & " " & vbCrLf
   End If
      
   Label_String = Label_String & "\1Bdw0928" & WardId & vbCrLf
   
   Label_String = Label_String & "\1Bq" & Format(CopyCount, "0###") & vbCrLf
   
End Function

Public Sub Label_FormFeed(Optional ByVal strOrdDiv As String = "L")
   
    Call GetBarInfo(strOrdDiv)
    Select Case strOrdDiv
        Case "A":
            objAPSbarcode.Label_FormFeed
        Case "B":
            objBBSbarcode.Label_FormFeed
        Case "L":
            objLISbarcode.Label_FormFeed
    End Select
   
'   Dim MyComm As Object
'   Dim StrX As String
'
'   StrX = Label_FeedString
'   If Not frmControls.MyComm.PortOpen Then Call Label_PortOpen
   
   'MyComm.CommPort = 2
   'MyComm.Settings = "9600,N,8,1"
   'MyComm.InputLen = 8192
   
   'MyComm.PortOpen = True
'   frmControls.MyComm.Output = StrX
   'MyComm.PortOpen = False
   
   'Set MyComm = Nothing

End Sub

Public Function Label_FeedString()

   Dim StrX As String
   
   StrX = ""
   StrX = StrX & "\1B@z" & vbCrLf
   StrX = StrX & "\1B@f09" & vbCrLf
   StrX = StrX & "\1Ba0901840208" & vbCrLf
   StrX = StrX & "\1Bf09" & vbCrLf
   StrX = StrX & "\1Bq0001" & vbCrLf
   
   Label_FeedString = StrX
   
End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
'  Barcode Type : Interleaved 2 of 5
'  Check Digit을 만들어 바코드 마지막에 추가한다.
'
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function AddCheckDigit(sBarcode As String) As String
    Dim iLen%
    Dim I%
    Dim iCheckSum%
    Dim iA%, iB%, iC%, iD%
    iLen = Len(sBarcode)
    iCheckSum = 0
    iA = 0
    iB = 0
    For I = 1 To iLen
        If I Mod 2 = 1 Then
            iB = iB + Val(Mid(sBarcode, I, 1))
        Else
            iA = iA + Val(Mid(sBarcode, I, 1))
        End If
    Next
    If iLen Mod 2 = 1 Then
        iC = iB * 3 + iA
    Else
        iC = iB + iA * 3
    End If
    iD = iC Mod 10
    If iD = 0 Then
        iCheckSum = 0
    Else
        iCheckSum = 10 - iD
    End If
    
    AddCheckDigit = sBarcode & Trim(str(iCheckSum))
End Function


'% Barcode Label 출력장수
'Public Property Let CopyCount(ByVal vData As Integer)
'    mvarCopyCount = vData
'End Property

'Public Property Get CopyCount() As Integer
'    CopyCount = mvarCopyCount
'End Property


'% 검사명
'Public Property Let TestNames(ByVal vData As String)
'    mvarTestNames = vData
'End Property

'Public Property Get TestNames() As String
'    TestNames = mvarTestNames
'End Property


'% 검체전달 Location
'Public Property Let Location(ByVal vData As String)
'    mvarLocation = vData
'End Property

'Public Property Get Location() As String
'    Location = mvarLocation
'End Property


'% 병동ID
'Public Property Let WardId(ByVal vData As String)
'    mvarWardId = vData
'End Property

'Public Property Get WardId() As String
'    WardId = mvarWardId
'End Property


'% 검체보관구분
'Public Property Let StoreCd(ByVal vData As String)
'    mvarStoreCd = vData
'End Property

'Public Property Get StoreCd() As String
'    StoreCd = mvarStoreCd
'End Property


'% 검체명
'Public Property Let SpcNm(ByVal vData As String)
'    mvarSpcNm = vData
'End Property

'Public Property Get SpcNm() As String
'    SpcNm = mvarSpcNm
'End Property


'% 환자명
'Public Property Let PtNm(ByVal vData As String)
'    mvarPtNm = vData
'End Property

'Public Property Get PtNm() As String
'    PtNm = mvarPtNm
'End Property


'% 환자ID
'Public Property Let PtId(ByVal vData As String)
'    mvarPtId = vData
'End Property

'Public Property Get PtId() As String
'    PtId = mvarPtId
'End Property


'% 검체번호
'Public Property Let SpcNo(ByVal vData As String)
'    mvarSpcNo = vData
'End Property

'Public Property Get SpcNo() As String
'    SpcNo = mvarSpcNo
'End Property


'% 응급여부
'Public Property Let StatFg(ByVal vData As String)
'    mvarStatFg = vData
'End Property

'Public Property Get StatFg() As String
'    StatFg = mvarStatFg
'End Property


'% 접수 Seq
'Public Property Let AccSeq(ByVal vData As String)
'    mvarAccSeq = vData
'End Property
'
'Public Property Get AccSeq() As String
'    AccSeq = mvarAccSeq
'End Property
'

'% 채혈일
'Public Property Let ColDt(ByVal vData As String)
'    mvarColDt = vData
'End Property
'
'Public Property Get ColDt() As String
'    ColDt = mvarColDt
'End Property


'% Work Area
'Public Property Let WorkArea(ByVal vData As String)
'    mvarWorkArea = vData
'End Property
'
'Public Property Get WorkArea() As String
'    WorkArea = mvarWorkArea
'End Property









