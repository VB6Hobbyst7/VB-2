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

'Private Const FormatNo = "09"
'
'Private Const LabelWidth = "0320"
'Private Const LabelLength = "0184"
'Private Const LabelTotLength = "0208"
'Private Const GapLength = 24
'Private Const PosYbar = "0035"
'Private Const PosY1 = "0020"
'Private Const PosY2 = "0019"
'Private Const PosY3 = "0055"
'Private Const PosY4 = "0120"
'Private Const PosY5 = "0125"
'Private Const PosY6 = "0140"
'Private Const PosY7 = "0150"
'Private Const PosY8 = "0165"
'Private Const PosY9 = "0035"
'Private Const PosYA = "0082"
'Private Const PosYB = "0100"
'
'Private Const PosXbld = "0150"       'building
'Private Const PosXbar = "0155"       'barcode
'Private Const PosXwa = "0040"        'workarea
'Private Const PosXdt = "0260"          'coldt
'Private Const PosXseq = "0370"       'accseq
'Private Const PosXsno = "0240"       'spc no
'Private Const PosXward = "0400"       'ward
'Private Const PosXstore = "0420"     'storecd
'Private Const PosXpnm = "0150"      'ptnm
'Private Const PosXpid = "0235"        'ptid
'Private Const PosXspc = "0340"       'spcnm
'Private Const PosXtest = "0150"       'testnm
'Private Const PosXstat = "0150"       'stat
'
'Private Const StFg = "00"
'Private Const FontDF = "0"
'Private Const FontSS = "1"
'Private Const FontMd = "2"
'Private Const FontLg = "3"
'Private Const FontKor1 = "0"  '바탕체
'Private Const FontKor2 = "1"  '굴림체
'Private Const NoRot = "00"
'Private Const Rot90 = "01"
'Private Const Rot180 = "02"
'Private Const Rot270 = "03"
'Private Const Reverse = "1"
'Private Const normal = "0"
'Private Const Code39 = "00"
'Private Const Code39C = "01"
'Private Const Code2of5 = "02"
'Private Const BarHeight = "0080"
'Private Const Readable = "01"
'Private Const NotReadable = "00"
'Private Const NarrowBar = "1"
'Private Const NWRatio = "1"
'Private Const BarLength = "12"
'Private Const Bold = "1"
'
'Private mvarWorkArea As String 'work area
'Private mvarColDt As String '채혈일
'Private mvarAccSeq As String 'accession sequence
'Private mvarStatFg As String '응급여부
'Private mvarSpcNo As String '검체번호
'Private mvarPtId As String '환자ID
'Private mvarPtNm As String '환자명
'Private mvarSpcNm As String '검체명
'Private mvarStoreCd As String '보관구분
'Private mvarWardId As String '병동ID
'Private mvarLocation As String '검체전달location
'Private mvarTestNames As String '검사명
'Private mvarCopyCount As Integer '출력장수


'Public Sub Clear()
'   mvarWorkArea = ""
'   mvarColDt = ""
'   mvarAccSeq = ""
'   mvarStatFg = ""
'   mvarSpcYY = ""
'   mvarSpcSeq = ""
'   mvarPtId = ""
'   mvarPtNm = ""
'   mvarSpcNm = ""
'   mvarStoreCd = ""
'   mvarWardId = ""
'   mvarLocation = ""
'   mvarTestNames = ""
'   mvarCopyCount = 0
'End Sub

'Public Sub GetBarInfo(ByVal strOrddiv As String)
'
'    '바코드 출력양식 읽어오기
'    Select Case strOrddiv
'        Case "A"
'            If Not blnAPSBarFg Then
'                Set objAPSbarcode.MyDb = DBConn
'                objAPSbarcode.ProjectCd = "APS"
'                Call objAPSbarcode.GetBarConfig
'                blnAPSBarFg = True
'            End If
'        Case "B"
'            If Not blnBBSBarFg Then
'                Set objBBSbarcode.MyDb = DBConn
'                objBBSbarcode.ProjectCd = "BBS"
'                Call objBBSbarcode.GetBarConfig
'                blnBBSBarFg = True
'            End If
'        Case "L"
'            If Not blnLISBarFg Then
'                Set objLISbarcode.MyDb = DBConn
'                objLISbarcode.ProjectCd = "LIS"
'                Call objLISbarcode.GetBarConfig
'                blnLISBarFg = True
'            End If
'    End Select
'
'End Sub


'Public Sub Label_PrintOut(ByVal strOrddiv As String, ByVal Location As Variant, ByVal WorkArea As Variant, _
'                          ByVal coldt As Variant, ByVal accseq As Variant, ByVal spcno As Variant, _
'                          ByVal ptid As Variant, ByVal ptnm As Variant, ByVal SpcNm As Variant, _
'                          ByVal StoreCd As Variant, ByVal StatFg As Variant, ByVal wardid As Variant, _
'                          ByVal orddt As Variant, ByVal ColTm As Variant, ByVal TestNames As Variant, _
'                          ByVal CopyCount As Variant, _
'                          Optional ByVal AccFg As Boolean = False, Optional ByVal FzFg As String = "")
'    Dim barString As String
'    Dim FileNo As Long
'    'Dim MyComm As Object
'    Dim PkSize As Integer
'
'    On Error GoTo SKIP
'
'    Call GetBarInfo(strOrddiv)
'
'    Select Case strOrddiv
'    Case "A":
'        Call objAPSbarcode.Label_PrintOut(Location, WorkArea, coldt, accseq, spcno, ptid, ptnm, _
'                        SpcNm, StoreCd, StatFg, wardid, orddt, ColTm, TestNames, CopyCount, AccFg, FzFg)
'    Case "B":
'        Call objBBSbarcode.Label_PrintOut(Location, WorkArea, coldt, accseq, spcno, ptid, ptnm, _
'                        SpcNm, StoreCd, StatFg, wardid, orddt, ColTm, TestNames, CopyCount, AccFg, FzFg)
'    Case "L"
'        Call objLISbarcode.Label_PrintOut(Location, WorkArea, coldt, accseq, spcno, ptid, ptnm, _
'                        SpcNm, StoreCd, StatFg, wardid, orddt, ColTm, TestNames, CopyCount, AccFg, FzFg)
'    End Select
'
'SKIP:
'
'End Sub

'Public Function Label_String(ByVal WorkArea As Variant, ByVal accdt As Variant, ByVal accseq As Variant, _
'                             ByVal StatFg As Variant, ByVal spcno As Variant, _
'                             ByVal ptid As Variant, ByVal ptnm As Variant, ByVal SpcNm As Variant, _
'                             ByVal StoreCd As Variant, ByVal wardid As Variant, ByVal Location As Variant, _
'                             ByVal TestNames As Variant, ByVal CopyCount As Variant, Optional ByVal AccFg As Boolean = False, _
'                             Optional ByVal orddt As String = "", Optional ByVal ColTm As String = "")
'
'
'
'   If CopyCount = 0 Then CopyCount = 1
'   If Len(TestNames) > 0 Then TestNames = Mid(TestNames, 1, Len(TestNames) - 1)
'   If AccFg Then accseq = accseq & Space(4 - Len(accseq)) & "V"
'   spcno = AddCheckDigit(CStr(spcno))    'check digit 추가
'
'   Label_String = ""
'   Label_String = Label_String & "\1B@z" & vbCrLf
'   Label_String = Label_String & "\1B@f" & FormatNo & vbCrLf
'   Label_String = Label_String & "\1Ba" & FormatNo & LabelLength & LabelTotLength & vbCrLf
'   Label_String = Label_String & "\1Bf" & FormatNo & vbCrLf
'
'   Label_String = Label_String & "\1Bbs" & FormatNo & "02" & StFg & PosXbar & PosYbar & BarLength & BarHeight & Code2of5 & NarrowBar & NWRatio & normal & NotReadable & vbCrLf    'Barcode Label
'
'   If wardid = "ER" And Location = "응급" Then
'        Label_String = Label_String & "\1Bds" & FormatNo & "02" & StFg & PosXbld & PosY1 & "04" & FontSS & FontMd & NoRot & Reverse & FontKor2 & Bold & vbCrLf   '건물
'   Else
'        Label_String = Label_String & "\1Bds" & FormatNo & "02" & StFg & PosXbld & PosY1 & "04" & FontSS & FontMd & NoRot & normal & FontKor2 & Bold & vbCrLf   '건물
'   End If
'   Label_String = Label_String & "\1Bds" & FormatNo & "04" & StFg & PosXwa & PosY1 & "02" & FontMd & FontMd & NoRot & normal & FontKor2 & normal & vbCrLf   'Workarea
'   Label_String = Label_String & "\1Bds" & FormatNo & "06" & StFg & PosXdt & PosY1 & "12" & FontSS & FontSS & NoRot & normal & FontKor2 & Bold & vbCrLf   'AccDt
'   'Label_String = Label_String & "\1Bds" & FormatNo & "06" & StFg & PosXdt & PosY1 & "05" & FontSS & FontSS & NoRot & Normal & FontKor2 & Bold & vbCRLF   '채혈일
'   Label_String = Label_String & "\1Bds" & FormatNo & "08" & StFg & PosXseq & PosY1 & "06" & FontMd & FontMd & NoRot & normal & FontKor2 & normal & vbCrLf   'AccSeq
'   Label_String = Label_String & "\1Bds" & FormatNo & "10" & StFg & PosXsno & PosY2 & "12" & FontSS & FontSS & NoRot & normal & FontKor2 & normal & vbCrLf   '검체번호
'   'Label_String = Label_String & "\1Bds" & FormatNo & "12" & StFg & PosXstore & PosY3 & "01" & FontMd & FontLg & NoRot & Normal & FontKor2 & Bold & vbCrLf   '보관구분
'   Label_String = Label_String & "\1Bds" & FormatNo & "14" & StFg & PosXward & PosYA & "05" & FontSS & FontSS & NoRot & normal & FontKor2 & normal & vbCrLf   '처방일
'   Label_String = Label_String & "\1Bds" & FormatNo & "16" & StFg & PosXward & PosYB & "05" & FontSS & FontSS & NoRot & normal & FontKor2 & normal & vbCrLf   '희망채혈일시
'   Label_String = Label_String & "\1Bds" & FormatNo & "18" & StFg & PosXpnm & PosY4 & "00" & FontSS & FontSS & NoRot & normal & FontKor2 & normal & vbCrLf   '환자명
'   Label_String = Label_String & "\1Bds" & FormatNo & "20" & StFg & PosXpid & PosY5 & "10" & FontSS & FontSS & NoRot & normal & FontKor2 & normal & vbCrLf   '환자ID
'   Label_String = Label_String & "\1Bds" & FormatNo & "22" & StFg & PosXspc & PosY4 & "10" & FontDF & FontSS & NoRot & normal & FontKor2 & Bold & vbCrLf   '검체명
'   Label_String = Label_String & "\1Bds" & FormatNo & "24" & StFg & PosXtest & PosY6 & "00" & FontSS & FontSS & NoRot & normal & FontKor2 & normal & vbCrLf   '검사명
'   Label_String = Label_String & "\1Bds" & FormatNo & "26" & StFg & PosXtest & PosY7 & "00" & FontSS & FontSS & NoRot & normal & FontKor2 & normal & vbCrLf   '검사명2
'
'   Label_String = Label_String & "\1Bds" & FormatNo & "28" & StFg & PosXward & PosY9 & "06" & FontSS & FontSS & NoRot & normal & FontKor2 & Bold & vbCrLf   'Ward Id
'
'   If Trim(StatFg) = "1" Then Label_String = Label_String & "\1Bls" & FormatNo & "02" & PosXstat & PosY8 & "0300" & "0000" & "0007" & vbCrLf   '응급
'
'   Label_String = Label_String & "\1Bbw0902" & spcno & vbCrLf
'
'   Label_String = Label_String & "\1Bdw0902" & Location & vbCrLf
'   Label_String = Label_String & "\1Bdw0904" & WorkArea & vbCrLf
'   Label_String = Label_String & "\1Bdw0906" & accdt & vbCrLf
'   'Label_String = Label_String & "\1Bdw0906" & ColDt & vbCRLF
'   Label_String = Label_String & "\1Bdw0908" & accseq & vbCrLf
'   Label_String = Label_String & "\1Bdw0910" & spcno & vbCrLf
'   'Label_String = Label_String & "\1Bdw0912" & StoreCd & vbCrLf
'   Label_String = Label_String & "\1Bdw0914" & orddt & vbCrLf
'   Label_String = Label_String & "\1Bdw0916" & ColTm & vbCrLf
'   Label_String = Label_String & "\1Bdw0918" & ptnm & vbCrLf
'   Label_String = Label_String & "\1Bdw0920" & ptid & vbCrLf
'   Label_String = Label_String & "\1Bdw0922" & SpcNm & vbCrLf
'
'   If Len(TestNames) > 36 Then
'      Label_String = Label_String & "\1Bdw0924" & Mid(TestNames, 1, 36) & vbCrLf
'      Label_String = Label_String & "\1Bdw0926" & Mid(TestNames, 37) & vbCrLf
'   Else
'      Label_String = Label_String & "\1Bdw0924" & TestNames & vbCrLf
'      Label_String = Label_String & "\1Bdw0926" & " " & vbCrLf
'   End If
'
'   Label_String = Label_String & "\1Bdw0928" & wardid & vbCrLf
'
'   Label_String = Label_String & "\1Bq" & Format(CopyCount, "0###") & vbCrLf
'
'End Function

'Public Sub Label_FormFeed(Optional ByVal strOrddiv As String = "L")
'
'    Call GetBarInfo(strOrddiv)
'    Select Case strOrddiv
'        Case "A":
'            objAPSbarcode.Label_FormFeed
'        Case "B":
'            objBBSbarcode.Label_FormFeed
'        Case "L":
'            objLISbarcode.Label_FormFeed
'    End Select
'
'End Sub
'
'Public Function Label_FeedString()
'
'   Dim StrX As String
'
'   StrX = ""
'   StrX = StrX & "\1B@z" & vbCrLf
'   StrX = StrX & "\1B@f09" & vbCrLf
'   StrX = StrX & "\1Ba0901840208" & vbCrLf
'   StrX = StrX & "\1Bf09" & vbCrLf
'   StrX = StrX & "\1Bq0001" & vbCrLf
'
'   Label_FeedString = StrX
'
'End Function

'Public Function BloodLabel_String(ByRef aryData() As Variant)
'
'    Dim i As Long
'
'    BloodLabel_String = ""
'    BloodLabel_String = BloodLabel_String & "\1B@z" & vbCrLf
'    BloodLabel_String = BloodLabel_String & "\1B@f" & FormatNo & vbCrLf
'    BloodLabel_String = BloodLabel_String & "\1Ba" & FormatNo & "0280" & "0296" & vbCrLf
'    BloodLabel_String = BloodLabel_String & "\1Bf" & FormatNo & vbCrLf
'
'
''    Label_String = Label_String & "\1Bds" & FormatNo & "04" & StFg & PosXwa & PosY1 & "02" & FontMd & FontMd & NoRot & normal & FontKor2 & normal & vbCrLf   'Workarea
'    BloodLabel_String = BloodLabel_String & "\1Bds" & FormatNo & "02" & StFg & "0050" & "0010" & "00" & "1" & "2" & "00" & "0" & FontKor2 & "1" & vbCrLf  'Workarea
'    BloodLabel_String = BloodLabel_String & "\1Bds" & FormatNo & "04" & StFg & "0220" & "0010" & "08" & "1" & "2" & NoRot & normal & FontKor2 & Bold & vbCrLf   'Workarea
'    BloodLabel_String = BloodLabel_String & "\1Bds" & FormatNo & "06" & StFg & "0030" & "0040" & "20" & "1" & "2" & NoRot & normal & FontKor2 & Bold & vbCrLf   'AccDt
'    BloodLabel_String = BloodLabel_String & "\1Bds" & FormatNo & "08" & StFg & "0030" & "0070" & "30" & "2" & "3" & NoRot & normal & FontKor2 & Bold & vbCrLf   'AccSeq
'    BloodLabel_String = BloodLabel_String & "\1Bds" & FormatNo & "10" & StFg & "0030" & "0110" & "20" & "1" & "2" & NoRot & normal & FontKor2 & Bold & vbCrLf   '검체번호
'
''    BloodLabel_String = BloodLabel_String & "\1Bds" & FormatNo & "12" & StFg & PosXstore & PosY3 & "01" & FontMd & FontLg & NoRot & normal & FontKor2 & Bold & vbCrLf   '보관구분
'    BloodLabel_String = BloodLabel_String & "\1Bds" & FormatNo & "14" & StFg & "0030" & "0165" & "15" & "0" & "1" & NoRot & normal & FontKor2 & Bold & vbCrLf   '처방일
'
'    BloodLabel_String = BloodLabel_String & "\1Bds" & FormatNo & "16" & StFg & "0030" & "0200" & "14" & "0" & FontSS & NoRot & normal & FontKor2 & Bold & vbCrLf   '희망채혈일시
'    BloodLabel_String = BloodLabel_String & "\1Bds" & FormatNo & "18" & StFg & "0270" & "0150" & "15" & FontSS & FontSS & NoRot & normal & FontKor2 & Bold & vbCrLf   '환자명
'    BloodLabel_String = BloodLabel_String & "\1Bds" & FormatNo & "20" & StFg & "0270" & "0180" & "00" & FontSS & FontSS & NoRot & normal & FontKor2 & normal & vbCrLf   '환자ID
'    BloodLabel_String = BloodLabel_String & "\1Bds" & FormatNo & "22" & StFg & "0025" & "0225" & "14" & "0" & "1" & NoRot & normal & FontKor2 & Bold & vbCrLf   '검체명
'    BloodLabel_String = BloodLabel_String & "\1Bds" & FormatNo & "24" & StFg & "0270" & "0210" & "00" & FontSS & FontSS & NoRot & normal & FontKor2 & normal & vbCrLf   '검사명
'    BloodLabel_String = BloodLabel_String & "\1Bds" & FormatNo & "26" & StFg & "0270" & "0240" & "00" & FontSS & FontSS & NoRot & normal & FontKor2 & normal & vbCrLf   '검사명2
'    BloodLabel_String = BloodLabel_String & "\1Bds" & FormatNo & "28" & StFg & "0350" & "0010" & "06" & "2" & "3" & NoRot & normal & FontKor2 & Bold & vbCrLf   'Ward Id
'
'    BloodLabel_String = BloodLabel_String & "\1Bds" & FormatNo & "30" & StFg & "0300" & "0010" & "06" & "2" & "3" & NoRot & normal & FontKor2 & Bold & vbCrLf   'Ward Id
'
'
'    BloodLabel_String = BloodLabel_String & "\1Bdw0902" & aryData(1) & vbCrLf
'    BloodLabel_String = BloodLabel_String & "\1Bdw0904" & aryData(2) & vbCrLf
'    BloodLabel_String = BloodLabel_String & "\1Bdw0906" & "진료과  : " & aryData(3) & vbCrLf
'    BloodLabel_String = BloodLabel_String & "\1Bdw0908" & "     " & aryData(6) & " " & aryData(4) & vbCrLf
'    BloodLabel_String = BloodLabel_String & "\1Bdw0910" & "혈액제제: " & aryData(13) & " " & aryData(7) & vbCrLf
''    BloodLabel_String = BloodLabel_String & "\1Bdw0912" & "AB+ " & vbCrLf
'
'    BloodLabel_String = BloodLabel_String & "\1Bdw0914" & "성별 : " & aryData(11) & vbCrLf
'
'    BloodLabel_String = BloodLabel_String & "\1Bdw0916" & aryData(10) & vbCrLf
'    BloodLabel_String = BloodLabel_String & "\1Bdw0918" & "혈액형  : " & aryData(5) & vbCrLf
'    BloodLabel_String = BloodLabel_String & "\1Bdw0920" & "검사자  : " & aryData(9) & vbCrLf
'    BloodLabel_String = BloodLabel_String & "\1Bdw0922" & "준비일:" & aryData(8) & vbCrLf
'    BloodLabel_String = BloodLabel_String & "\1Bdw0924" & "출고일시: " & vbCrLf
'    BloodLabel_String = BloodLabel_String & "\1Bdw0926" & "수령자  : " & vbCrLf
'
'    If aryData(12) = "1" Then
'        BloodLabel_String = BloodLabel_String & "\1Bdw0928" & "S" & vbCrLf
'    Else
'        BloodLabel_String = BloodLabel_String & "\1Bdw0928" & "" & vbCrLf
'    End If
'
'    If aryData(14) <> "" Then
'        BloodLabel_String = BloodLabel_String & "\1Bdw0930" & "R" & vbCrLf
'    Else
'        BloodLabel_String = BloodLabel_String & "\1Bdw0930" & "" & vbCrLf
'    End If
'
'    BloodLabel_String = BloodLabel_String & "\1Bq" & Format(1, "0###") & vbCrLf
'
'End Function



'
'
''* * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
''
''  Barcode Type : Interleaved 2 of 5
''  Check Digit을 만들어 바코드 마지막에 추가한다.
''
''* * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'Public Function AddCheckDigit(sBarcode As String) As String
'    Dim iLen%
'    Dim i%
'    Dim iCheckSum%
'    Dim iA%, iB%, iC%, id%
'    iLen = Len(sBarcode)
'    iCheckSum = 0
'    iA = 0
'    iB = 0
'    For i = 1 To iLen
'        If i Mod 2 = 1 Then
'            iB = iB + Val(Mid(sBarcode, i, 1))
'        Else
'            iA = iA + Val(Mid(sBarcode, i, 1))
'        End If
'    Next
'    If iLen Mod 2 = 1 Then
'        iC = iB * 3 + iA
'    Else
'        iC = iB + iA * 3
'    End If
'    id = iC Mod 10
'    If id = 0 Then
'        iCheckSum = 0
'    Else
'        iCheckSum = 10 - id
'    End If
'
'    AddCheckDigit = sBarcode & Trim(Str(iCheckSum))
'End Function

Option Explicit

Private mvarPort    As String
Private mvarKind    As String
Private mvarBarName As String

Private Sub Label_PortOpen()
    With frmControls.MyComm
        If .PortOpen Then Exit Sub
        .CommPort = mvarPort
        .Settings = "9600,N,8,1"
        .InputLen = 8192
        .Handshaking = comXOnXoff
        .InputMode = comInputModeText
        If Not .PortOpen Then .PortOpen = True
    End With
End Sub

Private Sub Label_PortClose()
    If frmControls.MyComm.PortOpen Then frmControls.MyComm.PortOpen = False
End Sub

Private Sub BarcodePrint_Argox(ByRef aryData() As Variant)
'--------------------------------------------------
' 감열지인 경우 위부분을 수정하라는 이야기이지요?
' D11=OD 로
'       .Output="OD" &vbcrlf
'--------------------------------------------------
'    aryData(1) = 환자ID:    aryData(2) = 환자명:     aryData(3) = 부서
'    aryData(4) = 혈액혈액형:aryData(5) = 환자혈액형: aryData(6) = 혈액번호:
'    aryData(7) = 용량:      aryData(8) = 보고일:     aryData(9) = 보고자:
'    aryData(10) = 주민번호: aryData(11) = 성별/나이 :aryData(12)=응급여부:
'    aryData(13)=제제:       aryData(14)=Irradiation
    
    

On Error GoTo Errors
    DoEvents
    With frmControls.MyComm
        .Output = "N" + vbCrLf
        .Output = "JF" + vbCrLf
        .Output = "OD" + vbCrLf
        .Output = "D10" & vbCrLf
        .Output = "S2" + vbCrLf
        .Output = "Q350,30" & vbCrLf
        .Output = "q550" & vbCrLf
        .Output = "ZB" + vbCrLf
        
        .Output = "A20,10,0,9,1,1,N," & ArgoxData("혈액번호:")
        .Output = "A20,40,0,9,1,1,N," & ArgoxData("혈액제제:")
        .Output = "A20,130,0,9,1,1,N," & ArgoxData("진료과:")
        .Output = "A20,100,0,9,1,1,N," & ArgoxData("결과:") 'ArgoxData("주민번호:")
        .Output = "A20,160,0,9,1,1,N," & ArgoxData("검사일:")
        .Output = "A20,190,0,9,1,1,N," & ArgoxData("검사자:")
        
        .Output = "A140,10,0,3,1,1,N," & ArgoxData(aryData(6))
        .Output = "A330,0,0,5,1,1,N," & ArgoxData(aryData(4))
        .Output = "A140,40,0,3,1,1,N," & ArgoxData(aryData(13) & "[" & aryData(7) & "]")
        .Output = "A20,70,0,3,1,1,N," & ArgoxData(aryData(1))
        .Output = "A140,70,0,9,1,1,N," & ArgoxData("[" & aryData(2) & "]")
        .Output = "A140,130,0,9,1,1,N," & ArgoxData(aryData(3))
        .Output = "A100,100,0,3,1,1,N," & ArgoxData(aryData(10)) '.Output = "A140,100,0,3,1,1,N," & ArgoxData(aryData(10))
        .Output = "A370,100,0,4,1,1,N," & ArgoxData(aryData(5)) '.Output = "A350,100,0,4,1,1,N," & ArgoxData(aryData(5))
        .Output = "A140,160,0,3,1,1,N," & ArgoxData(aryData(8))
        .Output = "A140,190,0,9,1,1,N," & ArgoxData(aryData(9))
        .Output = "A350,70,0,3,1,1,N," & ArgoxData(aryData(11))
        If aryData(12) = "1" Then
            .Output = "A350,130,0,9,1,1,R," & ArgoxData("응급")
        End If
        If aryData(14) = "1" Then
            .Output = "A350,190,0,9,1,1,R," & ArgoxData("Irr")
        End If
        .Output = "P1" + vbCrLf
    End With
    Exit Sub
Errors:
    MsgBox "바코드 출력오류입니다." & vbCrLf & _
           "바코드 종류와 포트를 확인하세요." & vbCrLf & _
           "현재설정된 기종은 " & mvarBarName & " 이며,설정 포트는 " & mvarPort & " 입니다.", vbInformation + vbOKOnly, "Info"
End Sub

Private Sub BarcodePrint_PD4(ByRef aryData() As Variant)
'--------------------------------------------------
' 감열지인 경우 위부분을 수정하라는 이야기이지요?
' D11=OD 로
'       .Output="OD" &vbcrlf
'--------------------------------------------------
'    aryData(1) = 환자ID:    aryData(2) = 환자명:     aryData(3) = 부서
'    aryData(4) = 혈액혈액형:aryData(5) = 환자혈액형: aryData(6) = 혈액번호:
'    aryData(7) = 용량:      aryData(8) = 보고일:     aryData(9) = 보고자:
'    aryData(10) = 주민번호: aryData(11) = 성별/나이 :aryData(12)=응급여부:
'    aryData(13)=제제:       aryData(14)=Irradiation
    
    

On Error GoTo Errors
    DoEvents
    With frmControls.MyComm
        .Output = "N" + vbCrLf
        .Output = "JF" + vbCrLf
        .Output = "OD" + vbCrLf
        .Output = "D10" & vbCrLf
        .Output = "S2" + vbCrLf
        .Output = "Q350,30" & vbCrLf 'Q280,24
        .Output = "q550" & vbCrLf 'q432
        .Output = "ZB" + vbCrLf
        
        .Output = "A20,10,0,8,1,1,N," & ArgoxData("혈액번호:")
        .Output = "A20,40,0,8,1,1,N," & ArgoxData("혈액제제:")
        .Output = "A20,130,0,8,1,1,N," & ArgoxData("진료과:")
        .Output = "A20,100,0,8,1,1,N," & ArgoxData("결과:") 'ArgoxData("주민번호:")
        .Output = "A20,160,0,8,1,1,N," & ArgoxData("검사일:")
        .Output = "A20,190,0,8,1,1,N," & ArgoxData("검사자:")
        
        .Output = "A140,10,0,3,1,1,N," & ArgoxData(aryData(6))
        .Output = "A330,0,0,5,1,1,N," & ArgoxData(aryData(4))
        .Output = "A140,40,0,3,1,1,N," & ArgoxData(aryData(13) & "[" & aryData(7) & "]")
        .Output = "A20,70,0,3,1,1,N," & ArgoxData(aryData(1))
        .Output = "A140,70,0,8,1,1,N," & ArgoxData("[" & aryData(2) & "]")
        .Output = "A140,130,0,8,1,1,N," & ArgoxData(aryData(3))
        .Output = "A100,100,0,3,1,1,N," & ArgoxData(aryData(10)) '.Output = "A140,100,0,3,1,1,N," & ArgoxData(aryData(10))
        .Output = "A370,100,0,4,1,1,N," & ArgoxData(aryData(5)) '.Output = "A350,100,0,4,1,1,N," & ArgoxData(aryData(5))
        .Output = "A140,160,0,3,1,1,N," & ArgoxData(aryData(8))
        .Output = "A140,190,0,8,1,1,N," & ArgoxData(aryData(9))
        .Output = "A350,70,0,3,1,1,N," & ArgoxData(aryData(11))
        If aryData(12) = "1" Then
            .Output = "A350,130,0,9,1,1,R," & ArgoxData("응급")
        End If
        If aryData(14) = "1" Then
            .Output = "A350,190,0,9,1,1,R," & ArgoxData("Irr")
        End If
        .Output = "P1" + vbCrLf
    End With
    Exit Sub
Errors:
    MsgBox "바코드 출력오류입니다." & vbCrLf & _
           "바코드 종류와 포트를 확인하세요." & vbCrLf & _
           "현재설정된 기종은 " & mvarBarName & " 이며,설정 포트는 " & mvarPort & " 입니다.", vbInformation + vbOKOnly, "Info"
End Sub

Private Function ArgoxData(ByVal sStr As String) As String
    ArgoxData = Chr(34) & sStr & Chr(34) & vbCrLf
End Function

Public Sub BloodLabel_Print(ByRef aryData() As Variant)

On Error GoTo Errors
    '바코드 정보를 읽어온다.
    Call GetBarCodeInfo
    '바코드 포트를 오픈한다.
    If Not frmControls.MyComm.PortOpen Then Label_PortOpen
    Select Case mvarKind
        Case 1
        Case 2
        Case 3: Call BarcodePrint_Argox(aryData())
        Case 4: Call BarcodePrint_PD4(aryData())
        Case 5
    End Select
    
    Call Label_PortClose
Errors:

End Sub

Private Sub GetBarCodeInfo()
    Dim strPath     As String
    
On Error GoTo Errors
    strPath = INIPath '"C:\Schweitzer\COMMON\DLL\BARCODE.INI"
    
    If Dir(strPath) = "" Then
        Call medSetINI("BAG", "KIND", "3", strPath)
        Call medSetINI("BAG", "PORT", "1", strPath)
    End If
    mvarKind = medGetINI("BAG", "KIND", strPath)
    mvarPort = medGetINI("BAG", "PORT", strPath)
    
    Select Case mvarKind
        Case 1: mvarBarName = "LEO60D"
        Case 2: mvarBarName = "Zebra T - 402"
        Case 3: mvarBarName = "Argox"
        Case 4: mvarBarName = "PD4"
        Case 5: mvarBarName = "Allegro"
    End Select
Errors:

End Sub

Public Sub PrintDonorLabel(ByRef aryData() As Variant)
'헌혈자 혈액 Tag 출력
On Error GoTo Errors
    '바코드 정보를 읽어온다.
    Call GetBarCodeInfo
    '바코드 포트를 오픈한다.
    If Not frmControls.MyComm.PortOpen Then Label_PortOpen
    Select Case mvarKind
        Case 1
        Case 2
        Case 3: Call PrintDonorLabel_Argox(aryData())
        Case 4: Call PrintDonorLabel_PD4(aryData())
        Case 5
    End Select
    
    Call Label_PortClose
Errors:

End Sub

Private Sub PrintDonorLabel_Argox(ByRef aryData() As Variant)
'aryData(1):혈액번호, aryData(2):혈액제제, aryData(3):용량
'aryData(4):혈액형, aryData(5):지정환자ID, aryData(6):환자명
'aryData(7):헌혈일, aryData(8):유효일, aryData(9):헌혈자
'aryData(10):헌혈자혈액형, aryData(11):바코드용 혈액번호

On Error GoTo Errors
    DoEvents
    With frmControls.MyComm
        .Output = "N" + vbCrLf
        .Output = "JF" + vbCrLf
        .Output = "OD" + vbCrLf
        .Output = "D10" & vbCrLf
        .Output = "S2" + vbCrLf
        .Output = "Q350,30" & vbCrLf 'Q280,24
        .Output = "q550" & vbCrLf 'q432
        .Output = "ZB" + vbCrLf
        
        .Output = "A15,10,0,9,1,1,N," & ArgoxData("혈액번호:")
        .Output = "A15,40,0,9,1,1,N," & ArgoxData("혈액제제:")
        .Output = "A15,70,0,9,1,1,N," & ArgoxData("지정환자:")
        .Output = "A15,100,0,9,1,1,N," & ArgoxData("채혈일:")
        .Output = "A225,100,0,9,1,1,N," & ArgoxData("유효일:")
        .Output = "A15,130,0,9,1,1,N," & ArgoxData("헌혈자:")
        
        .Output = "A135,10,0,3,1,1,N," & ArgoxData(aryData(1))
        .Output = "A135,40,0,3,1,1,N," & ArgoxData(aryData(2) & "[" & aryData(3) & "]")
        .Output = "A320,0,0,5,1,1,N," & ArgoxData(aryData(4))
        .Output = "A135,70,0,3,1,1,N," & ArgoxData(aryData(5))
        .Output = "A250,70,0,9,1,1,N," & ArgoxData("[" & aryData(6) & "]")
        .Output = "A110,100,0,3,1,1,N," & ArgoxData(aryData(7))
        .Output = "A315,100,0,3,1,1,N," & ArgoxData(aryData(8))
        .Output = "A140,130,0,9,1,1,N," & ArgoxData(aryData(9))
        .Output = "A300,130,0,3,1,1,N," & ArgoxData(aryData(10))
        'Bx축,y축,,코드39,,너비,높이,
        .Output = "B25,160,0,3,2,6,90,N," & ArgoxData(aryData(11)) '되는넘
        
        .Output = "P1" + vbCrLf
    End With
    Exit Sub
Errors:
    MsgBox "바코드 출력오류입니다." & vbCrLf & _
           "바코드 종류와 포트를 확인하세요." & vbCrLf & _
           "현재설정된 기종은 " & mvarBarName & " 이며,설정 포트는 " & mvarPort & " 입니다.", vbInformation + vbOKOnly, "Info"
End Sub

Private Sub PrintDonorLabel_PD4(ByRef aryData() As Variant)
'aryData(1):혈액번호, aryData(2):혈액제제, aryData(3):용량
'aryData(4):혈액형, aryData(5):지정환자ID, aryData(6):환자명
'aryData(7):헌혈일, aryData(8):유효일, aryData(9):헌혈자
'aryData(10):헌혈자혈액형, aryData(11):바코드용 혈액번호

On Error GoTo Errors
    DoEvents
    With frmControls.MyComm
        .Output = "N" + vbCrLf
        .Output = "JF" + vbCrLf
        .Output = "OD" + vbCrLf
        .Output = "D10" & vbCrLf
        .Output = "S2" + vbCrLf
        .Output = "Q350,30" & vbCrLf 'Q280,24
        .Output = "q550" & vbCrLf 'q432
        .Output = "ZB" + vbCrLf
        
        .Output = "A15,10,0,8,1,1,N," & ArgoxData("혈액번호:")
        .Output = "A15,40,0,8,1,1,N," & ArgoxData("혈액제제:")
        .Output = "A15,70,0,8,1,1,N," & ArgoxData("지정환자:")
        .Output = "A15,100,0,8,1,1,N," & ArgoxData("채혈일:")
        .Output = "A225,100,0,8,1,1,N," & ArgoxData("유효일:")
        .Output = "A15,130,0,8,1,1,N," & ArgoxData("헌혈자:")
        
        .Output = "A135,10,0,3,1,1,N," & ArgoxData(aryData(1))
        .Output = "A135,40,0,3,1,1,N," & ArgoxData(aryData(2) & "[" & aryData(3) & "]")
        .Output = "A320,0,0,5,1,1,N," & ArgoxData(aryData(4))
        .Output = "A135,70,0,3,1,1,N," & ArgoxData(aryData(5))
        .Output = "A250,70,0,8,1,1,N," & ArgoxData("[" & aryData(6) & "]")
        .Output = "A110,100,0,3,1,1,N," & ArgoxData(aryData(7))
        .Output = "A315,100,0,3,1,1,N," & ArgoxData(aryData(8))
        .Output = "A140,130,0,8,1,1,N," & ArgoxData(aryData(9))
        .Output = "A300,130,0,3,1,1,N," & ArgoxData(aryData(10))
        'Bx축,y축,,코드39,,너비,높이,
        .Output = "B25,160,0,3,2,6,90,N," & ArgoxData(aryData(11)) '되는넘
        
        .Output = "P1" + vbCrLf
    End With
    Exit Sub
Errors:
    MsgBox "바코드 출력오류입니다." & vbCrLf & _
           "바코드 종류와 포트를 확인하세요." & vbCrLf & _
           "현재설정된 기종은 " & mvarBarName & " 이며,설정 포트는 " & mvarPort & " 입니다.", vbInformation + vbOKOnly, "Info"
End Sub
