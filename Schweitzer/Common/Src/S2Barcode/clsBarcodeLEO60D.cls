VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'========================
'Leo60D(길병원)
'========================
'-----------------------------------------------------------------------------------
'  문자열
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
'  바코드
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
'  라인
'  Command ls : Syntax - '\1Bls' & p1 & p2 & p3 & p4 & p5 & p6 & p7
'                       p1 - Format No. (2 byte)
'                       p2 - Element No. (2 byte)
'                       p3 - ST-X (4 byte) : Horizontal start position (X) in dots. (0000~0447)
'                       p4 - ST-Y (4 byte) : Vertical start position (Y) in dots. (0000~2200)
'                       p5 - Horizontal length in dots (4 byte).
'                       p6 - Vertical length in dots (4 byte).
'                       p7 - Thick (4 byte) : 0000 ~ 0007
'-----------------------------------------------------------------------------------
'  종료(stream 의 끝)
'  Command q :  Syntax - '\1q" & p1
'                       p1 - Quantity (4 byte) : Number of copies of each label. (0001~9999)
'-----------------------------------------------------------------------------------

'==================
'Zebra (성모자애용)
'==================
'-----------------------------------------------------------------------------------
'  문자열출력

'  문자열예제     : ^XA
'                   ^FO 80,80^AE21,10^FD 출력문자열^FS
'                   ^XZ
'  Syntax 설명
'  Command ds : Syntax - ^XA : Operning Bracket으로 Format의 시작을 표시
'                        ^XZ : 바코드 문자 전송의 마지막:(XA와 한쌍)
'                        ^FO : 인쇄할항목의 (X축,Y축)을 정의(X,Y축dot수:0-9999)
'                        ^Ax,a,b,c : 1. x=폰트종류(A,B,C(D),E,F,G,H)
'                                    2. a=문자회전(기본값:N=Normal,R=90도,I=180도,B=270도)
'                                    3. b=문자높이,b=문자넓이
'                        ^FD : 인쇄할데이터를 정의
'                        ^FS : 인쇄할데이터의 마지막을 정의(^FD와 한쌍)

'-----------------------------------------------------------------------------------
'  바코드출력
'  바코드예제     : ^XA
'                   ^FO 80,80^AE21,10^FD 출력문자열^FS
'                   ^FO 80,160^BY2,3,100^Bx,,,,^FD 123456 ^FS
'                   ^XZ
'  Syntax -           ^BY a,b,c : 1. a=기존 Module(좁은바)의 넓이(기본값[2dot] ,범위[1-10dot])
'                                 2. b=넓은 Bar와 좁은Bar의 비율=Ratio(기본값[3:1],범위[2.0-3.0])
'                                 3. c=Bar의 높이(기본값[10Dot])
'                     ^Bx : x=적용가능한 Barcode의 정의
'                           x정의: 1. x=2(inTerleaved2 of 5)
'                                  2. x=3(code 39)
'                     *inTerleaved 2 of 5
'                     ex) B2 a,e,b,f,g
'                                 Check Disit의 사용(N=no,Y=yes :Default[N])
'                                바코드위의 문자인쇄여부(Default[N])
'                               바코드 밑의 문자인쇄여부(Default[Y])
'                              Bar의 높이(기본값: 10Dot)
'                             문자회전(기본값: N=Normal,R=90도,I=180도,B=270도)
'                     문자열) B2N,100,Y,N,N
'                             [DATA는 짝수개이어야 하며,홀수개의 DATA일경우는 자동으로 DATA의
'                             의 앞부분에 '0'(ZERO)가추가된다
'                     *CODE39
'                     ex) B2 a,e,b,f,g
'                                  바코드위의 문자인쇄여부(Default[N])
'                                바코드 밑의 문자인쇄여부(Default[Y])
'                               Bar의 높이(기본값: 10Dot)
'                              Module 43 Check Disit(N=no,Y=yes :Default[N])
'                             문자회전(기본값: N=Normal,R=90도,I=180도,B=270도)
'                     문자열) B3N,N,100,Y,N
'                             영문DATA는 대문자 이어야한다.
'                             '*'는 시작종료문자로 Printer에서 자동으로 붙혀준다.
'-----------------------------------------------------------------------------------
'==바코드속성지정==
'  *프린터속도 설정
'             ex) ^XA
'                 ^PRD
'                 ^FO 80,80^AE21,10^FD 출력문자열^FS
'                 ^XZ
'  Syntax - ^PRp,s : 1. p= Printer Speed => A(2),B(3),C(4 or 5),D(6),E(8)
'                    2. s= Slew Speed    => A(2),B(3),C(4 or 5),D(6),E(8)
'  *출력기준점 설정
'             ex)LHx,y : 1. x=x축의 Dot수(기본값:0 ,범위: 0-9999)
'                        2. y=y축의 Dot수(기본값:0 ,범위: 0-9999)
'                ^LH명령이 지정되면 프린터를 끄거나 계속하야 일어나는 Format를 바꿀때까지
'                 Setting된 값은 보존된다.
'  *Lable 길이 정의
'             ex)^LLy  : 1. y=Y축의 길이(단위:dot ,80dot=1Cm)
'  *Lable의 Data회전
'             ex)^FWR  : (기준점:N=Normal,R=90도,I=180도,B=270도)
'-----------------------------------------------------------------------------------
'==바코드 한글 출력==
'

Private Const CS_FotmatNo = "09"
Private Const CS_FotmatNo1 = "08"
Private Const CS_FotmatNo2 = "07"
Private Const CS_StFg = "00"
Private Const CS_FrozenFg = "(F)"

Private Const Readable = "01"
Private Const NotReadable = "00"
Private Const NarrowBar = "1"
Private Const NWRatio = "1"

'Private mobjMyDb            As Connection
Private objBarConfig        As New clsBarConfig
Private mvarObjbarH         As clsDictionary
Private objDicBAR           As clsDictionary
   
Private mudtBarData()       As tpBarData
Private mudtBarcode         As tpBarcode
Private mudtStatInfo        As tpStatInfo
Private mvarFontKor         As enFontKor
Private mvarRotation        As enRotation

Private mvarLabelWidth      As String
Private mvarLabelLength     As String
Private mvarLabelTotLength  As String
Private mvarAccCheck        As String
Private mvarBarKind         As String
Private mvarWorkArea        As String   'work area
Private mvarColDt           As String   '채혈일
Private mvarAccSeq          As String   'accession sequence
Private mvarStatFg          As String   '응급여부
Private mvarSpcNo           As String   '검체번호
Private mvarPtId            As String   '환자ID
Private mvarPtNm            As String   '환자명
Private mvarSpcNm           As String   '검체명
Private mvarStoreCd         As String   '보관구분
Private mvarWardId          As String   '병동ID
Private mvarLocation        As String   '검체전달location
Private mvarTestNames       As String   '검사명
Private mvarProjectCd       As String
Private mvarCopyCount       As Integer  '출력장수
Private mvarPortNo          As Long


'길병원 병동 메뉴 땜시 추가하였음.(kjg: 2002/02/20)
Private mvarWardTmp         As String

Public Property Let WardTmp(ByVal vData As String)
    mvarWardTmp = vData
End Property


Private Sub Class_Initialize()
'-- Initialize : Barcode출력 설정 내용을 읽어온다
    Set objDicBAR = New clsDictionary
    objDicBAR.Clear
    objDicBAR.FieldInialize "han", "hex,val"
    
    ReDim mudtBarData(14)
End Sub

'Public Property Set MyDb(ByVal vData As Connection)
''-- Database Object
'    Set mobjMyDb = vData
'    Set dbconn = vData
'End Property

Public Property Let ProjectCd(ByVal vData As String)
'-- 프로젝트 코드
    mvarProjectCd = vData
End Property
Public Property Get ProjectCd() As String
    ProjectCd = mvarProjectCd
End Property

Public Property Set TableInfo(ByVal vData As Object)
'-- 프로젝트 코드
    Set objTables = vData
End Property

Public Property Set FieldInfo(ByVal vData As Object)
    Set objFields = vData
End Property

'Public Property Let BarinfoTable(ByVal vData As String)
''-- 프로젝트 코드
'    T_COM002 = vData
'End Property

'Public Property Let BarHunTable(ByVal vData As String)
''-- 프로젝트 코드
'    T_COM004 = vData
'End Property

Friend Property Let StatInfo(ByRef vData As tpStatInfo)
'-- 응급여부 출력양식
    mudtStatInfo = vData
End Property
Friend Property Get StatInfo() As tpStatInfo
    StatInfo = mudtStatInfo
End Property

Friend Property Let Barcode(ByRef vData As tpBarcode)
'-- 바코드 출력양식
    mudtBarcode = vData
End Property
Friend Property Get Barcode() As tpBarcode
    Barcode = mudtBarcode
End Property

Friend Property Let BarData(ByVal iIndex As enBarData, ByRef vData As tpBarData)
'-- 각종 정보 출력양식
    mudtBarData(iIndex) = vData
End Property
Friend Property Get BarData(ByVal iIndex As enBarData) As tpBarData
    BarData = mudtBarData(iIndex)
End Property

Public Property Let LabelWidth(ByVal vData As String)
'-- 라벨 너비
    mvarLabelWidth = vData
End Property
Public Property Get LabelWidth() As String
    LabelWidth = mvarLabelWidth
End Property

Public Property Let LabelLength(ByVal vData As String)
'-- 라벨 길이
    mvarLabelLength = vData
End Property
Public Property Get LabelLength() As String
    LabelLength = mvarLabelLength
End Property

Public Property Let LabelTotLength(ByVal vData As String)
'-- 라벨 전체길이
    mvarLabelTotLength = vData
End Property
Public Property Get LabelTotLength() As String
    LabelTotLength = mvarLabelTotLength
End Property

Public Property Let PortNo(ByVal vData As Long)
'-- 시리얼 포트 번호
    mvarPortNo = vData
End Property
Public Property Get PortNo() As Long
    PortNo = mvarPortNo
End Property

Public Property Let AccCheck(ByVal vData As String)
'-- 접수번호 체크
    mvarAccCheck = vData
End Property
Public Property Get AccCheck() As String
    AccCheck = mvarAccCheck
End Property

Friend Property Let FontKor(ByVal vData As enFontKor)
'-- 한글 폰트
    mvarFontKor = vData
End Property
Friend Property Get FontKor() As enFontKor
    FontKor = mvarFontKor
End Property

Friend Property Let Rotation(ByVal vData As enRotation)
'-- Rotation
    mvarRotation = vData
End Property
Friend Property Get Rotation() As enRotation
    Rotation = mvarRotation
End Property
'% Barcode Label 출력장수
Public Property Let copycount(ByVal vData As Integer)
    mvarCopyCount = vData
End Property

Public Property Get copycount() As Integer
    copycount = mvarCopyCount
End Property
Public Property Let BarKind(ByVal vData As String)
'-- BarCode Kind
    mvarBarKind = vData
End Property
Public Property Get BarKind() As String
    BarKind = mvarBarKind
End Property

'% 검사명
Public Property Let TestNames(ByVal vData As String)
    mvarTestNames = vData
End Property

Public Property Get TestNames() As String
    TestNames = mvarTestNames
End Property


'% 검체전달 Location
Public Property Let Location(ByVal vData As String)
    mvarLocation = vData
End Property

Public Property Get Location() As String
    Location = mvarLocation
End Property


'% 병동ID
Public Property Let WardId(ByVal vData As String)
    mvarWardId = vData
End Property

Public Property Get WardId() As String
    WardId = mvarWardId
End Property


'% 검체보관구분
Public Property Let StoreCd(ByVal vData As String)
    mvarStoreCd = vData
End Property

Public Property Get StoreCd() As String
    StoreCd = mvarStoreCd
End Property


'% 검체명
Public Property Let SpcNm(ByVal vData As String)
    mvarSpcNm = vData
End Property

Public Property Get SpcNm() As String
    SpcNm = mvarSpcNm
End Property


'% 환자명
Public Property Let PtNm(ByVal vData As String)
    mvarPtNm = vData
End Property

Public Property Get PtNm() As String
    PtNm = mvarPtNm
End Property


'% 환자ID
Public Property Let Ptid(ByVal vData As String)
    mvarPtId = vData
End Property

Public Property Get Ptid() As String
    Ptid = mvarPtId
End Property


'% 검체번호
Public Property Let SpcNo(ByVal vData As String)
    mvarSpcNo = vData
End Property

Public Property Get SpcNo() As String
    SpcNo = mvarSpcNo
End Property


'% 응급여부
Public Property Let StatFg(ByVal vData As String)
    mvarStatFg = vData
End Property

Public Property Get StatFg() As String
    StatFg = mvarStatFg
End Property


'% 접수 Seq
Public Property Let AccSeq(ByVal vData As String)
    mvarAccSeq = vData
End Property

Public Property Get AccSeq() As String
    AccSeq = mvarAccSeq
End Property


'% 채혈일
Public Property Let ColDt(ByVal vData As String)
    mvarColDt = vData
End Property

Public Property Get ColDt() As String
    ColDt = mvarColDt
End Property


'% Work Area
Public Property Let WorkArea(ByVal vData As String)
    mvarWorkArea = vData
End Property

Public Property Get WorkArea() As String
    WorkArea = mvarWorkArea
End Property
'한글 이미지....
Public Property Set ObjBarH(ByVal vData As clsDictionary)
    Set mvarObjbarH = vData
End Property

Public Property Get ObjBarH() As clsDictionary
    Set ObjBarH = mvarObjbarH
End Property

Public Sub SetBarConfig()
'-- 바코드 출력양식 Setting 폼 로드
    frmBarConfig.ProjectId = mvarProjectCd
    frmBarConfig.Show
End Sub

Public Sub GetBarConfig()
    
    Call objBarConfig.GetConfig(mvarProjectCd, Me)

End Sub

Public Sub Get_PortNo()
    Call objBarConfig.GetPortNO(Me)
End Sub
Public Sub Clear()
    mvarWorkArea = ""
    mvarColDt = ""
    mvarAccSeq = ""
    mvarStatFg = ""
'    mvarSpcYY = ""
'    mvarSpcSeq = ""
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
'-- 시리얼 포트 오픈
    On Error GoTo ERROR

    With frmComm
        If .MyComm.PortOpen Then Exit Sub
        .MyComm.CommPort = mvarPortNo
        .MyComm.Settings = "9600,N,8,1"
        .MyComm.InputLen = 8192
        
        If Not .MyComm.PortOpen Then .MyComm.PortOpen = True
    End With
    Exit Sub
ERROR:
    MsgBox Err.Description
End Sub
Public Sub Label_FormFeed()
   
   Dim MyComm As Object
   Dim StrX As String

   StrX = Label_FeedString
   
   Call Label_PortOpen
   
   On Error GoTo ERROR
   
   frmComm.MyComm.Output = StrX

   Call Label_PortClose
   Exit Sub
ERROR:
    MsgBox Err.Description, vbCritical + vbOKOnly, "바코드출력에러"
End Sub

Public Function Label_FeedString()

    Dim StrX As String
   
'    StrX = Label_String("", "", "", "", "", "", "", "", "", "", "", "", "", "", 1)
'
    StrX = ""
    StrX = StrX & "\1B@z" & vbCrLf
    StrX = StrX & "\1B@f" & CS_FotmatNo2 & vbCrLf
    StrX = StrX & "\1Ba" & CS_FotmatNo2 & mvarLabelLength & mvarLabelTotLength & vbCrLf
    StrX = StrX & "\1Bf" & CS_FotmatNo2 & vbCrLf
    StrX = StrX & "\1Bq0001" & vbCrLf

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
    Dim i%
    Dim iCheckSum%
    Dim iA%, iB%, iC%, iD%
    iLen = Len(sBarcode)
    iCheckSum = 0
    iA = 0
    iB = 0
    For i = 1 To iLen
        If i Mod 2 = 1 Then
            iB = iB + Val(Mid(sBarcode, i, 1))
        Else
            iA = iA + Val(Mid(sBarcode, i, 1))
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
    
    AddCheckDigit = sBarcode & Trim(Str(iCheckSum))
End Function

Public Sub Label_PortClose()
'-- 시리얼 포트 닫기
    If frmComm.MyComm.PortOpen Then frmComm.MyComm.PortOpen = False

End Sub

Public Sub BloodLabel_PrintOut(ByVal PtNm As Variant, ByVal Ptid As Variant, ByVal WardId As Variant, _
                          ByVal CompNm As Variant, ByVal Volume As Variant, _
                          ByVal BldType As Variant, ByVal BldNo As Variant, ByVal TestDt As Variant, _
                          ByVal copycount As Variant)
'-- 라벨 출력
    Dim PkSize    As Integer
    Dim BarString As String
    
    Select Case BarKind
        Case "1"
                BarString = LEO_BloodLabel(PtNm, Ptid, WardId, CompNm, Volume, BldType, BldNo, TestDt, copycount)
                                
        Case "2"
                BarString = Zebra_BloodLabel(PtNm, Ptid, WardId, CompNm, Volume, BldType, BldNo, TestDt, copycount)
    End Select
    
    
    PkSize = 250
    
    On Error GoTo ERROR
    With frmComm
        
        Call Label_PortOpen
        
        If Len(BarString) > PkSize Then
            .MyComm.Output = Mid(BarString, 1, PkSize)
            While (Len(BarString)) > PkSize
                   BarString = Mid(BarString, PkSize + 1)
                   .MyComm.Output = Mid(BarString, 1, PkSize)
            Wend
        Else
            .MyComm.Output = BarString
        End If
        
        Call Label_PortClose
        
    End With
    Exit Sub
ERROR:
   MsgBox Err.Description, vbCritical + vbOKOnly, "바코드출력에러"
End Sub

Public Function LEO_BloodLabel(ByVal PtNm As Variant, ByVal Ptid As Variant, ByVal WardId As Variant, _
                                ByVal CompNm As Variant, ByVal Volume As Variant, ByVal BldType As Variant, _
                                ByVal BldNo As Variant, ByVal TestDt As Variant, ByVal copycount As Variant) As String
                                
    Dim i         As enBarData
    Dim BarString As String
    Dim strElemNo As String
    
    
    If copycount = 0 Then copycount = 1
    
    BarString = ""
    BarString = BarString & "\1B@z" & vbCrLf
    BarString = BarString & "\1B@f" & CS_FotmatNo1 & vbCrLf
    BarString = BarString & "\1Ba" & CS_FotmatNo1 & mvarLabelLength & mvarLabelTotLength & vbCrLf
    BarString = BarString & "\1Bf" & CS_FotmatNo1 & vbCrLf
    
    'Barcode Label
    If mudtBarcode.PrtFg = "1" Then
        BarString = BarString & "\1Bbs" & CS_FotmatNo1 & "02" & CS_StFg & _
                       mudtBarcode.PosX & mudtBarcode.PosY & mudtBarcode.Length & mudtBarcode.Height & _
                       Format(mudtBarcode.Style, "0#") & NarrowBar & NWRatio & mvarRotation & NotReadable & vbCrLf
    End If
    
    '기타 데이타
    For i = 1 To 8
        With mudtBarData(i)
            If .PrtFg = "1" Then
                strElemNo = Format(i * 2, "0#")
                .ElementNo = strElemNo
                BarString = BarString & "\1Bds" & CS_FotmatNo1 & strElemNo & CS_StFg & .PosX & .PosY & _
                            .Length & .FontX & .FontY & Format(mvarRotation, "0#") & .ReverseFg & mvarFontKor & _
                            .BoldFg & vbCrLf
            End If
        End With
    Next
    
    '바코드
    If mudtBarcode.PrtFg = "1" Then
        BarString = BarString & "\1Bbw" & CS_FotmatNo1 & "02" & SpcNo & vbCrLf
    End If
    
    '기타 데이타
    For i = 1 To 8
        With mudtBarData(i)
            If .PrtFg = "1" Then
                BarString = BarString & "\1Bdw" & CS_FotmatNo1 & .ElementNo
                Select Case .Key
                    Case DAT_BLDPTNM: BarString = BarString & PtNm & vbCrLf
                    Case DAT_BLDPTID: BarString = BarString & Ptid & vbCrLf
                    Case DAT_WARDID:  BarString = BarString & WardId & vbCrLf
                    Case DAT_COMPNM:  BarString = BarString & CompNm & vbCrLf
                    Case DAT_VOLUME:  BarString = BarString & Volume & vbCrLf
                    Case DAT_BLDTYPE: BarString = BarString & BldType & vbCrLf
                    Case DAT_BLDNO:   BarString = BarString & BldNo & vbCrLf
                    Case DAT_TESTDT:  BarString = BarString & TestDt & vbCrLf
                End Select
            End If
        End With
    Next
    
    
    LEO_BloodLabel = BarString & "\1Bq" & Format(copycount, "0###") & vbCrLf
    
End Function

Public Function Zebra_BloodLabel(ByVal PtNm As Variant, ByVal Ptid As Variant, ByVal WardId As Variant, _
                                  ByVal CompNm As Variant, ByVal Volume As Variant, ByVal BldType As Variant, _
                                  ByVal BldNo As Variant, ByVal TestDt As Variant, ByVal copycount As Variant) As String
'-- 라벨 출력
    Dim BarString   As String
    Dim i           As enBarData
    Dim mode        As String
    Dim FontX       As Long
    Dim FontY       As Long
    
    BarString = BarString & "^XA" & vbCrLf
   
    For i = 1 To 8
        With mudtBarData(i)
            If .PrtFg = "1" Then
                Select Case .FontX
                    Case "1": FontX = "15"
                    Case "2": FontX = "18"
                    Case "3": FontX = "20"
                End Select
                Select Case .FontY
                    Case "1": FontY = "15"
                    Case "2": FontY = "18"
                    Case "3": FontY = "20"
                End Select
                Select Case .Key
                    Case DAT_BLDPTNM: BarString = BarString & "^FO" & .PosX & "," & .PosY & _
                                                              "^AE" & FontX & "," & FontY & _
                                                              "^FD" & Get_HANGUL(PtNm, .PosX, .PosY)
                    Case DAT_BLDPTID: BarString = BarString & "^FO" & .PosX & "," & .PosY & _
                                                              "^AE" & FontX & "," & FontY & _
                                                              "^FD" & Ptid & "^FD" & vbCrLf
                    Case DAT_WARDID:  BarString = BarString & "^FO" & .PosX & "," & .PosY & _
                                                              "^AE" & FontX & "," & FontY & _
                                                              "^FD" & WardId & "^FS" & vbCrLf
                    Case DAT_COMPNM:  BarString = BarString & "^FO" & .PosX & "," & .PosY & _
                                                              "^AE" & FontX & "," & FontY & _
                                                              "^FD" & CompNm & "^FS" & vbCrLf
                    Case DAT_VOLUME:  BarString = BarString & "^FO" & .PosX & "," & .PosY & _
                                                              "^AE" & FontX & "," & FontY & _
                                                              "^FD" & Volume & "^FS" & vbCrLf
                    Case DAT_BLDTYPE: BarString = BarString & "^FO" & .PosX & "," & .PosY & _
                                                              "^AE" & FontX & "," & FontY & _
                                                              "^FD" & BldType & "^FS & vbCrLf"
                    Case DAT_BLDNO:   BarString = BarString & "^FO" & .PosX & "," & .PosY & _
                                                              "^AE" & FontX & "," & FontY & _
                                                              "^FD" & BldNo & "^FS" & vbCrLf
                    Case DAT_TESTDT:  BarString = BarString & "^FO" & .PosX & "," & .PosY & _
                                                              "^AE" & FontX & "," & FontY & _
                                                              "^FD" & TestDt & "^FS" & vbCrLf
                End Select
            End If
        End With
    Next
    
    '             "^FO40,50^BY3,4^B2N,90,N,N,N^FD123456^FS" & vbCrLf & _ '바코드
    '             "^FO15,50^BY2,2.5^B3N,N,90,N,N^FD123456^FS" & vbCrLf &
    
    If mudtBarcode.PrtFg = "1" Then     '바코드 출력
        '바코드 기종 결정
        mode = mudtBarcode.Style
        Select Case mode
            Case "1", "2"
                BarString = BarString & "^FO" & mudtBarcode.PosX & "," & mudtBarcode.PosY & "BY2,2.5" & "^B3N,N,90,N,N" & "^FD" & SpcNo & "^FS" & vbCrLf
            Case "3"
                BarString = BarString & "^FO" & mudtBarcode.PosX & "," & mudtBarcode.PosY & "BY3,4" & "^B2N,90,N,N,N" & "^FD" & SpcNo & "^FS" & vbCrLf
        End Select
    End If
    
        
    Zebra_BloodLabel = BarString & "^PQ" & copycount & vbCrLf & "^XZ"             '바코드 출력 문자열 완성

End Function

Public Function Gil_ER_BarCodeLabel(ByVal Location As Variant, ByVal WorkArea As Variant, ByVal AccDt As Variant, _
                          ByVal AccSeq As Variant, ByVal SpcNo As Variant, _
                          ByVal Ptid As Variant, ByVal PtNm As Variant, ByVal SpcNm As Variant, _
                          ByVal StoreCd As Variant, ByVal StatFg As Variant, ByVal WardId As Variant, _
                          ByVal OrdDt As Variant, ByVal ColTm As Variant, _
                          ByVal TestNames As Variant, ByVal copycount As Variant, _
                          Optional ByVal AccFg As Boolean = False, Optional ByVal FrozenFg As Boolean = False)
'-- 라벨 출력

    Dim BarString As String
    Dim PkSize As Integer
         
    If mvarWardTmp <> "" Then
        If Mid(SpcNo, 2, 1) = "0" Then
            Call objBarConfig.GetConfig("LIS", Me)
        Else
            Select Case Mid(SpcNo, 1, 1)
                Case 1: '임상
                     Call objBarConfig.GetConfig("LIS", Me)
                Case 2: '해부
                     Call objBarConfig.GetConfig("APS", Me)
                Case 3: '혈액
                     Call objBarConfig.GetConfig("BBS", Me)
            End Select
        End If
    Else
        Select Case Mid(SpcNo, 1, 1)
            Case 1: '임상
                     Call objBarConfig.GetConfig("LIS", Me)
            Case 2: '해부
                     Call objBarConfig.GetConfig("APS", Me)
            Case 3: '혈액
                     Call objBarConfig.GetConfig("BBS", Me)
        End Select
    End If
    
    Gil_ER_BarCodeLabel = Label_String(Location, WorkArea, AccDt, AccSeq, SpcNo, Ptid, PtNm, SpcNm, StoreCd, _
                                       StatFg, WardId, OrdDt, ColTm, TestNames, copycount, AccFg, FrozenFg)
     
    
End Function

Public Sub Label_PrintOut(ByVal Location As Variant, ByVal WorkArea As Variant, ByVal AccDt As Variant, _
                          ByVal AccSeq As Variant, ByVal SpcNo As Variant, _
                          ByVal Ptid As Variant, ByVal PtNm As Variant, ByVal SpcNm As Variant, _
                          ByVal StoreCd As Variant, ByVal StatFg As Variant, ByVal WardId As Variant, _
                          ByVal OrdDt As Variant, ByVal ColTm As Variant, _
                          ByVal TestNames As Variant, ByVal copycount As Variant, _
                          Optional ByVal AccFg As Boolean = False, Optional ByVal FrozenFg As Boolean = False)
'-- 라벨 출력

    Dim BarString   As String
    Dim strBarKey   As String
    Dim strBarKind  As String
    Dim PkSize      As Integer
    
    Select Case Mid(SpcNo, 1, 1)
        Case 1: strBarKey = "LIS" '임상
        Case 2: strBarKey = "APS" '해부
        Case 3: strBarKey = "BBS" '혈액
    End Select
    
    strBarKind = SQL_BARKIND(strBarKey)
         
    If mvarWardTmp <> "" Then
        If Mid(SpcNo, 2, 1) = "0" Then
            strBarKey = "LIS"
        End If
    End If
    
    strBarKey = strBarKey & COL_DIV & strBarKind
    
    '바코드에 대한 기본정보 가져오기-----------------
    Call objBarConfig.GetConfig(strBarKey, Me)
    
    Select Case mvarBarKind
        Case "1"    'LEO60D
            PtNm = Mid(PtNm, 1, 3)
            BarString = Label_String(Location, WorkArea, AccDt, AccSeq, SpcNo, Ptid, PtNm, SpcNm, StoreCd, _
                                     StatFg, WardId, OrdDt, ColTm, TestNames, copycount, AccFg, FrozenFg)
        Case "2"    'Zebra
            PtNm = Mid(PtNm, 1, 3)
            BarString = Zebra_Label_String(Location, WorkArea, AccDt, AccSeq, SpcNo, Ptid, PtNm, SpcNm, StoreCd, _
                                           StatFg, WardId, OrdDt, ColTm, TestNames, copycount, AccFg, FrozenFg)
        Case "3"    'Argox
            Call ArgoxPrint(Location, WorkArea, AccDt, AccSeq, SpcNo, Ptid, PtNm, SpcNm, StoreCd, StatFg, WardId, OrdDt, ColTm, TestNames, copycount, AccFg, FrozenFg)
            Exit Sub
        Case "4"    'PD4
            Call PD4Print(Location, WorkArea, AccDt, AccSeq, SpcNo, Ptid, PtNm, SpcNm, StoreCd, StatFg, WardId, OrdDt, ColTm, TestNames, copycount, AccFg, FrozenFg)
            Exit Sub
        Case "5"    'Allegro
            Call AllegroPrint(Location, WorkArea, AccDt, AccSeq, SpcNo, Ptid, PtNm, SpcNm, StoreCd, StatFg, WardId, OrdDt, ColTm, TestNames, copycount, AccFg, FrozenFg)
            Exit Sub
        Case "6"    'Citizen
            '## 5.0.4: 이상대(2005-02-15)
            '   - Citizen T4 바코드장비 추가
            Call CitizenPrint(Location, WorkArea, AccDt, AccSeq, SpcNo, Ptid, PtNm, SpcNm, StoreCd, StatFg, WardId, OrdDt, ColTm, TestNames, copycount, AccFg, FrozenFg)
            Exit Sub
    End Select
    
    PkSize = 250
    On Error GoTo ERROR
    With frmComm
         Call Label_PortOpen
         If Len(BarString) > PkSize Then
             .MyComm.Output = Mid(BarString, 1, PkSize)
             While (Len(BarString)) > PkSize
                    BarString = Mid(BarString, PkSize + 1)
                    .MyComm.Output = Mid(BarString, 1, PkSize)
             Wend
         Else
             .MyComm.Output = BarString
         End If
         
         Call Label_PortClose
         
    End With
    Exit Sub
ERROR:
    MsgBox Err.Description, vbCritical + vbOKOnly, "바코드출력에러"
End Sub
Public Function Zebra_Label_String(ByVal Location As Variant, ByVal WorkArea As Variant, _
                                   ByVal AccDt As Variant, ByVal AccSeq As Variant, _
                                   ByVal SpcNo As Variant, ByVal Ptid As Variant, _
                                   ByVal PtNm As Variant, ByVal SpcNm As Variant, _
                                   ByVal StoreCd As Variant, ByVal StatFg As Variant, _
                                   ByVal WardId As Variant, ByVal OrdDt As Variant, _
                                   ByVal ColTm As Variant, ByVal TestNames As Variant, _
                                   ByVal copycount As Variant, Optional ByVal AccFg As Boolean = False, _
                                   Optional ByVal FrozenFg As Boolean = False) As String
' ----------------------------
' 바코드 기종: Zebra(성모자애)
' 라벨출력문자열 생성
' ----------------------------
    
    Dim i           As enBarData
    Dim strElemNo   As String
    Dim strTestNm1  As String
    Dim strTestNm2  As String
    Dim strFrozenFg As String
    
    Dim strLabel    As String       '라벨 출력용 문자열 생성
    Dim mode        As String       '바코드 기종
    Dim FontX       As String
    Dim FontY       As String
    Dim StrSpc      As String
    
    Dim STRTMP      As String
    Dim fp          As Integer
    
    fp = FreeFile
    On Error Resume Next
    
    Open "C:\PRN" For Output As fp
    
    If copycount = 0 Then copycount = 1
    If Len(TestNames) > 0 Then
        If Mid(TestNames, Len(TestNames), 1) = "," Then _
            TestNames = Mid(TestNames, 1, Len(TestNames) - 1)
    End If
    If (mvarAccCheck = "1") And AccFg Then AccSeq = AccSeq & Space(4 - Len(AccSeq)) & "V"
    If FrozenFg Then strFrozenFg = CS_FrozenFg
    
    StrSpc = Mid(SpcNo, 1, 2) & "-" & Format(Mid(SpcNo, 3), "#########")
    SpcNo = AddCheckDigit(CStr(SpcNo))    'check digit 추가
    
    strTestNm1 = Mid(TestNames, 1, 36) & " "
    strTestNm2 = Mid(TestNames, 37, 36) & " "
    
    Select Case Mid(SpcNo, 1, 1)
        Case 1: '임상
            STRTMP = "^XA" & vbCrLf
            'WORKAREA
            STRTMP = STRTMP & "^FO0060,0025^A020,20^FD" & "(" & WorkArea & ")^FS" & vbCrLf
            'ACCDT
            STRTMP = STRTMP & "^FO0110,0030^A015,15^FD" & AccDt & "^FS" & vbCrLf
            'ACCSEQ
            STRTMP = STRTMP & "^FO0190,0030^A020,20^FD" & AccSeq & "^FS" & vbCrLf
            'SPCNO
            STRTMP = STRTMP & "^FO0080,0045^A015,15^FD" & SpcNo & "^FS" & vbCrLf
            '
            STRTMP = STRTMP & "^FO0260,0060^A015,15^FD" & WardId & "^FS" & vbCrLf
            '
            STRTMP = STRTMP & "^FO0260,0087^A015,15^FD" & OrdDt & "^FS" & vbCrLf
            '
            STRTMP = STRTMP & "^FO0260,0102^A015,15^FD" & ColTm & "^FS" & vbCrLf
            '임상
            STRTMP = STRTMP & "~DGNAME1" + ",0072,003," & "0000000FF00E1C380E300C0E700E0E60060E60060E60060E60060E700E0E300C0E1C380E0FF00E0000000000000000007FFFFE7FFFFE70000E70000E70000E7FFFFE7FFFFE000000" & vbCrLf
            STRTMP = STRTMP & "~DGNAME2" + ",0072,003," & "00000003C03803C03803C03803C03803C03803C03E07E03E07E0380FF0381E78383C3C38781E380000000000000FFFE01FFFF03800383000183000183800381FFFF00FFFE0000000" & vbCrLf
            STRTMP = STRTMP & "^FO0010,0025^XGNAME1,1,1" + "^FS" & vbCrLf
            STRTMP = STRTMP & "^FO0035,0025^XGNAME2,1,1" + "^FS" & vbCrLf
            
            STRTMP = STRTMP & Get_HANGUL(PtNm, "110", "145")
            
            STRTMP = STRTMP & "^FO0020,0150^A015,15^FD" & Ptid & "^FS" & vbCrLf
            
            STRTMP = STRTMP & "^FO0240,0150^A015,15^FD^FD" & SpcNm & "^FS" & vbCrLf
            'TESTNM1
            STRTMP = STRTMP & Get_HANGUL(strTestNm1, "0000", "0167")
'            STRTMP = STRTMP & "^FO0000,0167^A015,15^FD" & strTestNm1 & "^FS" & vbCrLf
            
            If Trim(strTestNm2) <> "" Then STRTMP = STRTMP & Get_HANGUL(strTestNm1, "0000", "0182")
            
            '초기
            'STRTMP = STRTMP & "^FO^FO0020,0060^B32,2^B2N,80,N,N,N^FD" & SpcNo & "^FS" & vbCrLf
            '2 OF 5
            STRTMP = STRTMP & "^FO^FO0020,0060^B2N,80,N,N,N^FD" & SpcNo & "^FS" & vbCrLf
'            'CODE39
'            STRTMP = STRTMP & "^FO^FO0020,0060^B3N,N,80,Y,N^FD" & SpcNo & "^FS" & vbCrLf
            
            STRTMP = STRTMP & "^PQ" & copycount & vbCrLf
            STRTMP = STRTMP & "^XZ"
        Case 2: '해부
            STRTMP = "^XA" & vbCrLf
            STRTMP = STRTMP & "^FO0060,0025^A020,20^FD" & "(AP)^FS" & vbCrLf
            STRTMP = STRTMP & "^FO0085,0045^A015,15^FD" & SpcNo & "^FS" & vbCrLf
            STRTMP = STRTMP & "^FO0260,0060^A015,15^FD" & WardId & "^FS" & vbCrLf
            STRTMP = STRTMP & "^FO0260,0087^A015,15^FD" & OrdDt & "^FS" & vbCrLf
            STRTMP = STRTMP & "^FO0260,0102^A015,15^FD" & ColTm & "^FS" & vbCrLf
            STRTMP = STRTMP & "~DGNAME1" + ",0072,003," & "0000000FE1CE0001CE7FFDCE0001CE0001CE07C1CE1FF1CE3C79CE701DCE701DFE701DFE701DCE701DCE701DCE701DCE701DCE701DCE701DCE3C79CE1FF1CE07C1CE0001CE000000" & vbCrLf
            STRTMP = STRTMP & "~DGNAME2" + ",0072,003," & "00000070000E70000E70000E70000E7FFFFE7FFFFE70000E70000E7FFFFE7FFFFE0000000000000000007FFFFE7FFFFE001C00001C00001C00001C00001C00001C00001C00000000" & vbCrLf
            STRTMP = STRTMP & "^FO0010,0025^XGNAME1,1,1" + "^FS" & vbCrLf
            STRTMP = STRTMP & "^FO0035,0025^XGNAME2,1,1" + "^FS" & vbCrLf
            STRTMP = STRTMP & Get_HANGUL(PtNm, "110", "145")
            STRTMP = STRTMP & "^FO0020,0150^A015,15^FD" & Ptid & "^FS" & vbCrLf
            STRTMP = STRTMP & "^FO0240,0150^A015,15^FD^FD" & SpcNm & "^FS" & vbCrLf
            STRTMP = STRTMP & Get_HANGUL(strTestNm1, "0000", "0167")
            If Trim(strTestNm2) <> "" Then STRTMP = STRTMP & Get_HANGUL(strTestNm1, "0000", "0182")
            '초기
            'STRTMP = STRTMP & "^FO^FO0020,0060^B32,2^B2N,80,N,N,N^FD" & SpcNo & "^FS" & vbCrLf
            '2 OF 5
            STRTMP = STRTMP & "^FO^FO0020,0060^B2N,80,N,N,N^FD" & SpcNo & "^FS" & vbCrLf
'            'CODE39
'            STRTMP = STRTMP & "^FO^FO0020,0060^B3N,N,80,Y,N^FD" & SpcNo & "^FS" & vbCrLf
            
            If FrozenFg = True Then STRTMP = STRTMP & "^FO0240,0030^A025,25^FD(F)^FS" & vbCrLf
            STRTMP = STRTMP & "^PQ" & copycount & vbCrLf
            STRTMP = STRTMP & "^XZ"
        Case 3: '혈액
            STRTMP = STRTMP & "^XA" & vbCrLf
            STRTMP = STRTMP & "^FO0060,0025^A025,25^FD(XM)^FS" & vbCrLf
            STRTMP = STRTMP & "^FO0150,0030^A025,25^FD" & AccSeq & "^FS" & vbCrLf
            STRTMP = STRTMP & "^FO0090,0050^A015,15^FD" & SpcNo & "^FS" & vbCrLf
            STRTMP = STRTMP & "^FO0260,0070^A015,15^FD" & WardId & "^FS" & vbCrLf
            STRTMP = STRTMP & "^FO0260,0100^A015,15^FD" & OrdDt & "^FS" & vbCrLf
            STRTMP = STRTMP & "^FO0260,0115^A015,15^FD" & ColTm & "^FS" & vbCrLf
            STRTMP = STRTMP & "~~~DGNAME1,0072,003,0000001FF80E00000E7FFE0E0000FE1FF8FE3C3C0E700E0E700EFE700EFE700E0E3C3C0E1FF80E0000007FFFFE7FFFFE00000E00000E7FFFFE7FFFFE7000007000007FFFFE7FFFFE" & vbCrLf
            STRTMP = STRTMP & "~~~DGNAME2,0072,003,0000000FC1CE1CE1CE3031CE7039CE6019CE6019FE6019FE6019CE7039CE3031CE1CE1CE0FC1CE0000000000000000007FFFFE7FFFFE00000E00000E00000E00000E00000E000000" & vbCrLf
            STRTMP = STRTMP & "^FO0010,0025^XGNAME1,1,1^FS" & vbCrLf
            STRTMP = STRTMP & "^FO0035,0025^XGNAME2,1,1^FS" & vbCrLf
            STRTMP = STRTMP & Get_HANGUL(PtNm, "180", "150")
            STRTMP = STRTMP & "^FO0020,0155^A015,15^FD" & Ptid & "^FS" & vbCrLf
            '초기
'            STRTMP = STRTMP & "^FO^FO0020,0065^B32,2^B2N,80,N,N,N^FD" & SpcNo & "^FS" & vbCrLf
            '2 OF 5
            STRTMP = STRTMP & "^FO^FO0020,0065^B2N,80,N,N,N^FD" & SpcNo & "^FS" & vbCrLf
'            'CODE39
'            STRTMP = STRTMP & "^FO^FO0020,0060^B3N,N,80,Y,N^FD" & SpcNo & "^FS" & vbCrLf
            
            STRTMP = STRTMP & "^PQ" & copycount & vbCrLf
            STRTMP = STRTMP & "^XZ"
        End Select
    Print #fp, STRTMP

    
    Close fp

End Function
Public Function Get_HANGUL(ByVal strHan As String, _
                           ByVal PosX As String, ByVal PosY As String) As String
    
    Dim Rs      As New Recordset
    Dim strSql  As String
    Dim strHOne As String       '하나씩 자른 문자열
    Dim strNew  As String       '새롭게 생성된 문자열
    Dim HLen    As Long         '찍는 한글을 길이
    Dim LPosX   As Long         '넘어온 X좌표위치
    Dim i       As Integer
    
    HLen = Len(strHan)
    
    LPosX = PosX
   ' ObjBarH
    For i = 1 To HLen Step 1
        strHOne = ""
        strHOne = Mid(strHan, i, 1)
        '-------------------------
        '한글, 영문, 숫자인지 체크
        '-------------------------
        If Abs(Asc(strHOne)) < 128 Then
            '-------------------------------------------------
            '영문, 숫자는 기냥 출력
            '-------------------------------------------------
            strNew = strNew & "^FO" & LPosX & "," & PosY & _
                              "^A0" & 15 & "," & 15 & _
                              "^FD" & strHOne & "^FS" & vbCrLf
            
            LPosX = LPosX + 9
        Else
            '---------------------------------------------------
            '한글인 경우는 이미지를 바코더에 다운로드 한후 출력
            '이미 Dictionary에 담아져 있는지를 찾아서 있으면
            'Dictionary에 값을 가지고오고 없으면 DB에서 불러온다
            '---------------------------------------------------
            If objDicBAR.Exists(strHOne) Then
                objDicBAR.KeyChange strHOne
                    strNew = strNew & "~DG" & objDicBAR.Fields("val") & ",00072,003," & _
                                              objDicBAR.Fields("hex") & vbCrLf
                    strNew = strNew & "^FO" & LPosX & "," & PosY & _
                                      "^XG" & objDicBAR.Fields("val") & ",1,1^FS" & vbCrLf
            Else
                strSql = " SELECT * from " & objTables.T_COM004 & _
                         " WHERE " & DBW("cdindex", "zebra", 2) & _
                         " AND " & DBW("cdval1", strHOne, 2)
                Rs.Open strSql, DBConn
                
                If Rs.EOF Then
'                If Rs.DBerror Then
'                    DBConn.DisplayErrors
                Else
                    'field1: Hex지정값, text1: Hex값
                    objDicBAR.AddNew Rs.Fields("cdval1").Value & "", _
                                   Join(Array(Rs.Fields("text1").Value & "", _
                                   Rs.Fields("field1").Value & ""), COL_DIV)
                    '바코드 다운로드
                    strNew = strNew & "~DG" & Rs.Fields("field1").Value & ",00072,003," & _
                                              Rs.Fields("text1").Value & "" & vbCrLf
                    '출력 문자열
                    strNew = strNew & "^FO" & LPosX & "," & PosY & _
                                      "^XG" & Rs.Fields("field1").Value & "" & ",1,1^FS" & vbCrLf
                End If
'                Rs.RsClose
                Set Rs = Nothing
            End If
            LPosX = LPosX + 25
        End If
    Next
    Get_HANGUL = strNew
    
End Function

Public Function ArgoxIntermecString(ByVal pString As String) As String
    ArgoxIntermecString = Chr(34) & pString & Chr(34) & vbLf
End Function

Public Function ArgoxPrint(ByVal pLocation As String, ByVal pWorkArea As String, ByVal pAccDt As String, _
                           ByVal pAccSeq As String, ByVal pSpcNo As String, ByVal pPtid As String, _
                           ByVal pPtNm As String, ByVal pSpcNm As String, ByVal pStoreCd As String, _
                           ByVal pStatFg As String, ByVal pWardId As String, ByVal pOrdDt As String, _
                           ByVal pColTm As String, ByVal pTestNames As String, ByVal pCopyCnt As String, _
                           Optional ByVal pAccFg As Boolean = False, Optional ByVal pFrozenFg As Boolean = False)
'-- 바코드 출력
    Dim strTestNm1  As String
    Dim strTestNm2  As String
    Dim strQValue1  As String
    Dim strQValue2  As String
    Dim strqValue   As String
    Dim strFont     As String
    Dim strFrozenFg As String
    Dim strPosition As String
    Dim strBarFor   As String
    Dim strDiv      As String
    Dim strTmpSpcno As String
    
    Dim ii          As Integer
    
'-- 시리얼 포트 오픈
On Error GoTo Errors

    With frmComm
        If Not .MyComm.PortOpen Then
            .MyComm.CommPort = mvarPortNo
            .MyComm.Settings = "9600,N,8,1"
            .MyComm.InputLen = 0
            .MyComm.Handshaking = comXOnXoff
            .MyComm.InputMode = comInputModeText
        End If
        If Not .MyComm.PortOpen Then .MyComm.PortOpen = True
    End With
    
    strDiv = "L"
    If Mid(pSpcNo, 1, 1) <> "1" Then strDiv = "B"
    
    If (mvarAccCheck = "1") And pAccFg Then pAccSeq = pAccSeq & " V"   ' 접수여부체크
    
    If pFrozenFg Then strFrozenFg = CS_FrozenFg
    
    If Len(pTestNames) > 0 Then
        If Mid(pTestNames, Len(pTestNames), 1) = "," Then pTestNames = Mid(pTestNames, 1, Len(pTestNames) - 1)
    End If
    
    
    If Len(pTestNames) > 20 Then
        If Mid(Mid(pTestNames, 1, 20), 20) = "," Then strTestNm1 = Mid(pTestNames, 1, 19)
        
        strTestNm1 = Mid(pTestNames, 1, 20) & "..." '30
    Else
        strTestNm1 = Mid(pTestNames, 1, 20)  '30
    End If
    
    strTestNm2 = Mid(pTestNames, 31)

    strQValue1 = Val(mvarLabelLength) * 0.8
    strQValue2 = (Val(mvarLabelTotLength) - Val(mvarLabelLength)) * 0.8
    strqValue = Val(mvarLabelWidth) * 0.8
    
    DoEvents
'    Call medSleep(1600)
    With frmComm.MyComm
        .Output = "N" + vbCrLf                                  '클리어 버퍼
        .Output = "JF" + vbCrLf                                 '빽피드
        .Output = "OD" + vbCrLf                                ' 감열 사용 옵션
        .Output = "D10" & vbCrLf                               'Density 기본값은 10 값이 높을수록 진하게
        .Output = "S2" + vbCrLf                                 '스피드 2 적정값
        .Output = "Q" & strQValue1 & "," & strQValue2 & vbCrLf  'Q값 3.5*0.8*100=280[라벨의 길이], 0.8*0.3cm(겝)=24
        .Output = "q" & strqValue & vbCrLf                      '5.4*0.8*100=432[라벨의 Width]
        .Output = "ZB" + vbCrLf                                 '프린트 이미지
        
        'Barcode Label
        Select Case mudtBarcode.Style
            Case BAR_CODE2OF5
                If mudtBarcode.Length = 11 Then '체크 디지트 사용하지 않는 경우
                    strBarFor = "2"
                ElseIf mudtBarcode.Length = 12 Then '체크 디지트 사용하는 경우
                    strBarFor = "2C"
                End If
            Case BAR_CODE39
                If mudtBarcode.Length = 11 Then '체크 디지트 사용하지 않는 경우
                    strBarFor = "3"
                ElseIf mudtBarcode.Length = 12 Then '체크 디지트 사용하는 경우
                    strBarFor = "3C"
                End If
        End Select
        
        If mudtBarcode.PrtFg = "1" Then
            .Output = "B" & Format(mudtBarcode.PosX, "#") & "," & Format(mudtBarcode.PosY, "#") & ",0," & strBarFor & ",2,6," & Format(mudtBarcode.Height, "#") & ",N," & _
                      ArgoxIntermecString(pSpcNo) 'ArgoxIntermecString(IIf(mudtBarcode.Length = 11, pSpcNo, AddCheckDigit(pSpcNo))) '
        End If

        If pStatFg = "1" Then
            .Output = "A" & Format(mudtStatInfo.PosX, "#") & "," & Format(mudtStatInfo.PosY, "#") & ",0,9,1,1,R," & ArgoxIntermecString("응급")
        End If
        
        '기타 데이타
        For ii = 1 To 14
            If mudtBarData(ii).PrtFg = "1" Then
                '한글인경우 '0'을 '9'로 변경
                strFont = mudtBarData(ii).FontX
                If strFont = "0" Then strFont = "9"
                strPosition = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N,"
                Select Case mudtBarData(ii).Key
                    Case DAT_LOCATION:
                            .Output = strPosition & ArgoxIntermecString(pLocation & pWorkArea)                                          '
                    Case DAT_WORKAREA:
                            .Output = strPosition & ArgoxIntermecString(pLocation & pWorkArea)
                    Case DAT_ACCDT:
                            .Output = strPosition & ArgoxIntermecString(pAccDt)
                    Case DAT_ACCSEQ:
                            .Output = strPosition & ArgoxIntermecString(pAccSeq)
                    Case DAT_SPCNO:
                            If strDiv = "L" Then
                                strTmpSpcno = pPtid & " " & Mid(pSpcNo, 1, 2) & "-" & Format(Mid(pSpcNo, 3), "#")
                                .Output = strPosition & ArgoxIntermecString(strTmpSpcno)
                            Else
                                .Output = strPosition & ArgoxIntermecString(pSpcNo)
                            End If
                    Case DAT_DEPT:
                            .Output = strPosition & ArgoxIntermecString(pWardId)
                    Case DAT_ORDDT:
                            .Output = strPosition & ArgoxIntermecString(pOrdDt)
                    Case DAT_COLTM:
                            .Output = strPosition & ArgoxIntermecString(pColTm)
                    Case DAT_PTNM:
                            .Output = strPosition & ArgoxIntermecString(pPtNm)
'                            .Output = "A10,140,0,9,1,1,N," & ArgoxIntermecString("정일환")
                    Case DAT_PTID:
                            If strDiv <> "L" Then
                                .Output = strPosition & ArgoxIntermecString(pPtid)
                            End If
                    Case DAT_SPCNM:
                            If pSpcNm = "전용용기" Then
'                            A0235,0120,0,2,1,1,N,"전용용기"
                                .Output = "A0235,0120,0,9,1,1,N," & ArgoxIntermecString(pSpcNm)
                            Else
                                .Output = strPosition & ArgoxIntermecString(pSpcNm)
                            End If
                    Case DAT_ORDNM1:
                            .Output = strPosition & ArgoxIntermecString(strTestNm1)
                    Case DAT_ORDNM2:
                            .Output = strPosition & ArgoxIntermecString(strTestNm2)
                    Case DAT_FROZEN:
                            .Output = strPosition & ArgoxIntermecString(strFrozenFg)
                End Select
            End If
            
        Next
        
        .Output = "P" & pCopyCnt & vbCrLf
        
        ' -- 시리얼포트 Close
        If .PortOpen Then .PortOpen = False
    End With
Errors:
End Function

Public Function PD4Print(ByVal pLocation As String, ByVal pWorkArea As String, ByVal pAccDt As String, _
                         ByVal pAccSeq As String, ByVal pSpcNo As String, ByVal pPtid As String, _
                         ByVal pPtNm As String, ByVal pSpcNm As String, ByVal pStoreCd As String, _
                         ByVal pStatFg As String, ByVal pWardId As String, ByVal pOrdDt As String, _
                         ByVal pColTm As String, ByVal pTestNames As String, ByVal pCopyCnt As String, _
                         Optional ByVal pAccFg As Boolean = False, Optional ByVal pFrozenFg As Boolean = False)
'-- 바코드 출력
    Dim strTestNm1  As String
    Dim strTestNm2  As String
    Dim strQValue1  As String
    Dim strQValue2  As String
    Dim strqValue   As String
    Dim strFont     As String
    Dim strFrozenFg As String
    Dim ii          As Integer
    Dim strDiv As String
    Dim strTmpSpcno As String
    Dim strBarFor As String
    
'-- 시리얼 포트 오픈
On Error GoTo Errors

    With frmComm
        If Not .MyComm.PortOpen Then
            .MyComm.CommPort = mvarPortNo
            .MyComm.Settings = "9600,N,8,1"
            .MyComm.InputLen = 0
            .MyComm.Handshaking = comXOnXoff
            .MyComm.InputMode = comInputModeText
        End If
        If Not .MyComm.PortOpen Then .MyComm.PortOpen = True
    End With
    
    strDiv = "L"
    If Mid(pSpcNo, 1, 1) <> "1" Then strDiv = "B"
    
    If (mvarAccCheck = "1") And pAccFg Then pAccSeq = pAccSeq & " V"   ' 접수여부체크
    
    If pFrozenFg Then strFrozenFg = CS_FrozenFg
    
    If Len(pTestNames) > 0 Then
        If Mid(pTestNames, Len(pTestNames), 1) = "," Then pTestNames = Mid(pTestNames, 1, Len(pTestNames) - 1)
    End If
    
    
    If Len(pTestNames) > 20 Then
        If Mid(Mid(pTestNames, 1, 20), 20) = "," Then strTestNm1 = Mid(pTestNames, 1, 19)
        
        strTestNm1 = Mid(pTestNames, 1, 20) & "..." '30
    Else
        strTestNm1 = Mid(pTestNames, 1, 20)  '30
    End If
'    strTestNm1 = Mid(pTestNames, 1, 30)
    strTestNm2 = Mid(pTestNames, 31)
    
    strQValue1 = Val(mvarLabelLength) * 0.8
    strQValue2 = (Val(mvarLabelTotLength) - Val(mvarLabelLength)) * 0.8
    strqValue = Val(mvarLabelWidth) * 0.8
    
    With frmComm.MyComm
        .Output = "N" + vbCrLf                                  '클리어 버퍼
        .Output = "JF" + vbCrLf                                 '빽피드
        .Output = "OD" + vbCrLf                                ' 감열 사용 옵션
        .Output = "D10" & vbCrLf                               'Density 기본값은 10 값이 높을수록 진하게
        .Output = "S2" + vbCrLf                                 '스피드 2 적정값
        .Output = "Q" & strQValue1 & "," & strQValue2 & vbCrLf  'Q값 3.5*0.8*100=280[라벨의 길이], 0.8*0.3cm(겝)=24
        .Output = "q" & strqValue & vbCrLf                      '5.4*0.8*100=432[라벨의 Width]
        .Output = "ZB" + vbCrLf                                 '프린트 이미지
    
'        .Output = "ON,D" & vbCrLf                                 '라벨 taken sensor 안하는거 Gap단위로
'        .Output = "Q" & strQValue1 & "," & strQValue2 & vbCrLf  'Q값 3.5*0.8*100=280[라벨의 길이], 0.8*0.3cm(겝)=24
'        .Output = "q" & strqValue & vbCrLf                      '5.4*0.8*100=432[라벨의 Width]
'        .Output = "S3" & vbCrLf                                 '스피드 3 적정값
'        .Output = "D15" & vbCrLf                                '진하기 조절
'        .Output = "JF" & vbCrLf                                 '빽피드
'        .Output = "ZB" & vbCrLf                                 '프린트 이미지
'        .Output = "N" & vbCrLf                                  '클리어 버퍼
        Select Case mudtBarcode.Style
            Case BAR_CODE2OF5
                If mudtBarcode.Length = 11 Then '체크 디지트 사용하지 않는 경우
                    strBarFor = "2"
                ElseIf mudtBarcode.Length = 12 Then '체크 디지트 사용하는 경우
                    strBarFor = "2C"
                End If
            Case BAR_CODE39
                If mudtBarcode.Length = 11 Then '체크 디지트 사용하지 않는 경우
                    strBarFor = "3"
                ElseIf mudtBarcode.Length = 12 Then '체크 디지트 사용하는 경우
                    strBarFor = "3C"
                End If
        End Select
        
        'Barcode Label
        If mudtBarcode.PrtFg = "1" Then '"B25,40,0,2C,2,6,75,N,""14000000238"""
            .Output = "B" & Format(mudtBarcode.PosX, "#") & "," & Format(mudtBarcode.PosY, "#") & ",0," & strBarFor & ",2,6," & Format(mudtBarcode.Height, "#") & ",N," & _
                      ArgoxIntermecString(pSpcNo) 'ArgoxIntermecString(IIf(mudtBarcode.Length = 11, pSpcNo, AddCheckDigit(pSpcNo)))
        End If
        
        If pStatFg = "1" Then '"A0230,0142,0,8,1,1,N,""응급"""
            .Output = "A" & Format(mudtStatInfo.PosX, "#") & "," & Format(mudtStatInfo.PosY, "#") & ",0,8,1,1,R," & ArgoxIntermecString("응급")
        End If
        
        '기타 데이타
        For ii = 1 To 14
            If mudtBarData(ii).PrtFg = "1" Then
                '한글인경우 '0'을 '8'로 변경
                strFont = mudtBarData(ii).FontX
                If strFont = "0" Then strFont = "8" '"9"
                Select Case mudtBarData(ii).Key
'--------------------------
'아르곡스와 통일하기 위해서 아르곡스 쏘스 그대로 사용함
'                    Case DAT_LOCATION:
'                            .Output = strPosition & ArgoxIntermecString(pLocation & pWorkArea)                                          '
                    Case DAT_WORKAREA: '"A0005,0010,0,8,1,1,N,""02임상""" & vbLf //한글로 출력
                            .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0,8,1,1,N," & _
                                      ArgoxIntermecString(pLocation & pWorkArea)
                    Case DAT_ACCDT: '"A0140,0010,0,1,1,1,N,""040429""" & vbLf
                            .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N," & _
                                      ArgoxIntermecString(pAccDt)
                    Case DAT_ACCSEQ: '"A0230,0008,0,2,1,1,N,""3 V""" & vbLf
                            .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N," & _
                                      ArgoxIntermecString(pAccSeq)
                    Case DAT_SPCNO: '"A0100,0025,0,1,1,1,N,""00000002 14-238""" & vbLf
                            If strDiv = "L" Then
                                strTmpSpcno = pPtid & " " & Mid(pSpcNo, 1, 2) & "-" & Format(Mid(pSpcNo, 3), "#")
                                .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N," & _
                                      ArgoxIntermecString(strTmpSpcno)
                            Else
                                .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N," & _
                                      ArgoxIntermecString(pSpcNo)
                            End If
                    Case DAT_DEPT: '"A0100,0010,0,1,1,1,N,""A10""" & vbLf '출력할때 영문으로 출력되므로 한글명 부서는 사용할 수 없음
                            .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N," & _
                                      ArgoxIntermecString(pWardId)
                    Case DAT_ORDDT: '사용안함
                            .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N," & _
                                      ArgoxIntermecString(pOrdDt)
                    Case DAT_COLTM: '사용안함
                            .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N," & _
                                     ArgoxIntermecString(pColTm)
                    Case DAT_PTNM: '"A0010,0140,0,8,1,1,N,""Endeuryu""" & vbLf //한글로 출력
                            .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0,8,1,1,N," & _
                                     ArgoxIntermecString(pPtNm)
                    Case DAT_PTID: '"A0100,0025,0,1,1,1,N,""00000002 14-238""" & vbLf '검체번호 찍을때 같이 찍어줌
                            If strDiv <> "L" Then
                                .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N," & _
                                          ArgoxIntermecString(pPtid)
                            End If
                    Case DAT_SPCNM: '"A0235,0120,0,2,1,1,N,""WB(E)""" & vbLf '영문 출력으로 한글로된 검체명 사용할 수 없음
                            .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N," & _
                                      ArgoxIntermecString(pSpcNm)
                    Case DAT_ORDNM1: '"A0010,0127,0,1,1,1,N,""CBC,Diff""" & vbLf '영문 출력으로 한글로된 검사명 사용할 수 없음
                            .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N," & _
                                      ArgoxIntermecString(strTestNm1)
                    Case DAT_ORDNM2: '사용안함
                            .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N," & _
                                      ArgoxIntermecString(strTestNm1)
                    Case DAT_FROZEN: '사용안함
                            .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N," & _
                                      ArgoxIntermecString(strTestNm1)
'--------------------------
'기존 쏘스
'                    Case DAT_LOCATION: '사용안함(workarea와 같이 찍어줌)
'                            .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N," & _
'                                      ArgoxIntermecString(pLocation & pWorkArea) & vbCrLf
'                    Case DAT_WORKAREA: '0005,0010
'                            .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",4," & strFont & ",1,1,N," & _
'                                      ArgoxIntermecString(pLocation & pWorkArea) '& vbLf
'                    Case DAT_ACCDT: '0140,0010
'                            .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N," & _
'                                      ArgoxIntermecString(pAccDt) '& vbLf
'                    Case DAT_ACCSEQ: '0230,0008
'                            .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N," & _
'                                      ArgoxIntermecString(pAccSeq) '& vbCrLf
'                    Case DAT_SPCNO: '사용안함(환자 아디와 같이 찍어줌)
'                            .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N," & _
'                                      ArgoxIntermecString(pSpcNo) '& vbCrLf
'                    Case DAT_DEPT: '0100,0010
'                            .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N," & _
'                                      ArgoxIntermecString("Part : " & pWardId) '& vbCrLf
'                    Case DAT_ORDDT: '사용안함
'                            .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N," & _
'                                      ArgoxIntermecString(pOrdDt) '& vbCrLf
'                    Case DAT_COLTM: '사용안함
'                            .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N," & _
'                                      ArgoxIntermecString("Time : " & pOrdDt & " " & pColTm) '& vbCrLf
'                    Case DAT_PTNM: '0010,0140
'                            .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N," & _
'                                      ArgoxIntermecString(pPtNm) '& vbCrLf
'                    Case DAT_PTID:
'                            .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N," & _
'                                      ArgoxIntermecString(pPtid) '& vbCrLf
'                    Case DAT_SPCNM:
'                            .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N," & _
'                                      ArgoxIntermecString(pSpcNm) '& vbCrLf
'                    Case DAT_ORDNM1:
'                            .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N," & _
'                                      ArgoxIntermecString("Item : " & strTestNm1) '& vbCrLf
'                    Case DAT_ORDNM2:
'                            .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N," & _
'                                      ArgoxIntermecString("       " & strTestNm2) '& vbCrLf
'                    Case DAT_FROZEN:
'                            .Output = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N," & _
'                                      ArgoxIntermecString(strFrozenFg) '& vbCrLf
                End Select
            End If
            
        Next
        
        DoEvents
        .Output = "P" & pCopyCnt & vbLf
        ' -- 시리얼포트 Close
        If .PortOpen Then .PortOpen = False
    End With
Errors:

End Function

Public Function AllegroPrint(ByVal pLocation As String, ByVal pWorkArea As String, ByVal pAccDt As String, _
                             ByVal pAccSeq As String, ByVal pSpcNo As String, ByVal pPtid As String, _
                             ByVal pPtNm As String, ByVal pSpcNm As String, ByVal pStoreCd As String, _
                             ByVal pStatFg As String, ByVal pWardId As String, ByVal pOrdDt As String, _
                             ByVal pColTm As String, ByVal pTestNames As String, ByVal pCopyCnt As String, _
                             Optional ByVal pAccFg As Boolean = False, Optional ByVal pFrozenFg As Boolean = False)
'-- 바코드 출력
    Dim strTestNm1  As String
    Dim strTestNm2  As String

    Dim strFrozenFg As String
    Dim strBuffer   As String
    Dim strBold     As String
    Dim strFontX    As String
    Dim strFontY    As String
    Dim strPosX     As String
    Dim strPosY     As String
    Dim ii          As Integer
    Dim jj          As Integer
    Dim kk          As Integer
    
'-- 시리얼 포트 오픈
On Error GoTo Errors

    With frmComm
        If Not .MyComm.PortOpen Then
            .MyComm.CommPort = mvarPortNo
            .MyComm.Settings = "9600,N,8,1"
            .MyComm.InputLen = 0
            .MyComm.Handshaking = comXOnXoff
            .MyComm.InputMode = comInputModeText
        End If
        If Not .MyComm.PortOpen Then .MyComm.PortOpen = True
    End With
    
    If (mvarAccCheck = "1") And pAccFg Then pAccSeq = pAccSeq & " V"   ' 접수여부체크
    pAccSeq = Trim(pAccSeq)
    
    If pFrozenFg Then strFrozenFg = CS_FrozenFg
    
    If Len(pTestNames) > 0 Then
        If Mid(pTestNames, Len(pTestNames), 1) = "," Then pTestNames = Mid(pTestNames, 1, Len(pTestNames) - 1)
    End If
    
    strTestNm1 = Mid(pTestNames, 1, 30)
    strTestNm2 = Mid(pTestNames, 31)
    

    
    With frmComm.MyComm
        strBuffer = Chr$(2) & "f260" & Chr$(13)
        strBuffer = strBuffer & Chr$(2) & "L" & Chr$(13)
        strBuffer = strBuffer & "D11" & Chr$(13)
       
        If mudtBarcode.PrtFg = "1" Then
            If mudtBarcode.Style = BAR_CODE39 Then
                strBuffer = strBuffer & "1d62" & Format(mudtBarcode.Height, "000") & Format(mudtBarcode.PosY, "0000") & Format(mudtBarcode.PosX, "0000") & pSpcNo & Chr$(13)
            ElseIf mudtBarcode.Style = BAR_CODE2OF5 Then
                strBuffer = strBuffer & "1a62" & Format(mudtBarcode.Height, "000") & Format(mudtBarcode.PosY, "0000") & Format(mudtBarcode.PosX, "0000") & pSpcNo & Chr$(13)
            End If
        End If
        
        If pStatFg = "1" Then
            strBuffer = strBuffer & "1Y" & mudtStatInfo.Width & Format(mudtStatInfo.PosY, "0000") & Format(mudtStatInfo.PosX, "0000") & "응" & Chr$(13)
            strBuffer = strBuffer & "1Y" & mudtStatInfo.Width & Format(mudtStatInfo.PosY, "0000") & Format(Val(mudtStatInfo.PosX) + 10, "0000") & "급" & Chr$(13)
        End If
        
        '기타 데이타
        For ii = 1 To 14
            If mudtBarData(ii).PrtFg = "1" Then
                '한글인경우 '0'을 '9'로 변경
                strBold = IIf(mudtBarData(ii).BoldFg = SET_BOLD, "2", "1")
                strFontX = mudtBarData(ii).FontX
                strFontY = mudtBarData(ii).FontY & "000"
                strPosX = Format(mudtBarData(ii).PosX, "0000")
                strPosY = Format(mudtBarData(ii).PosY, "0000")
                
                Select Case mudtBarData(ii).Key
                    Case DAT_LOCATION:
                            kk = strPosX
                            For jj = 1 To Len(pLocation)
                                strBuffer = strBuffer & "1Y" & strFontX & strFontY & strPosY & Format$(kk, "0000")
                                strBuffer = strBuffer & Mid(pLocation, jj, 1) & Chr$(13)
                                 kk = kk + 10
                            Next
                    Case DAT_WORKAREA:
                            strBuffer = strBuffer & "1" & strBold & strFontX & strFontY & strPosY & strPosX & pWorkArea & Chr$(13)
                    Case DAT_ACCDT:
                            strBuffer = strBuffer & "1" & strBold & strFontX & strFontY & strPosY & strPosX & pAccDt & Chr$(13)
                    Case DAT_ACCSEQ:
                            strBuffer = strBuffer & "1" & strBold & strFontX & strFontY & strPosY & strPosX & pAccSeq & Chr$(13)
                    Case DAT_SPCNO:
                            strBuffer = strBuffer & "1" & strBold & strFontX & strFontY & strPosY & strPosX & pSpcNo & Chr$(13)
                    Case DAT_DEPT:
                            strBuffer = strBuffer & "1" & strBold & strFontX & strFontY & strPosY & strPosX & pWardId & Chr$(13)
                    Case DAT_ORDDT:
                            strBuffer = strBuffer & "1" & strBold & strFontX & strFontY & strPosY & strPosX & pOrdDt & Chr$(13)
                    Case DAT_COLTM:
                            strBuffer = strBuffer & "1" & strBold & strFontX & strFontY & strPosY & strPosX & pColTm & Chr$(13)
                    Case DAT_PTNM:
                            kk = strPosX
                            For jj = 1 To Len(pPtNm)
                                strBuffer = strBuffer & "1Y" & strFontX & strFontY & strPosY & Format$(kk, "0000")
                                strBuffer = strBuffer & Mid(pPtNm, jj, 1) & Chr$(13)
                                kk = kk + 10
                            Next
                    Case DAT_PTID:
                            strBuffer = strBuffer & "1" & strBold & strFontX & strFontY & strPosY & strPosX & pPtid & Chr$(13)
                    Case DAT_SPCNM:
                            strBuffer = strBuffer & "1" & strBold & strFontX & strFontY & strPosY & strPosX & pSpcNm & Chr$(13)
                    Case DAT_ORDNM1:
                            strBuffer = strBuffer & "1" & strBold & strFontX & strFontY & strPosY & strPosX & strTestNm1 & Chr$(13)
                    Case DAT_ORDNM2:
                            strBuffer = strBuffer & "1" & strBold & strFontX & strFontY & strPosY & strPosX & strTestNm2 & Chr$(13)
                    Case DAT_FROZEN:
                            strBuffer = strBuffer & "1" & strBold & strFontX & strFontY & strPosY & strPosX & strFrozenFg & Chr$(13)
                End Select
            End If
        Next
        strBuffer = strBuffer & "E" & Chr$(13)
        .Output = strBuffer
        ' -- 시리얼포트 Close
        If .PortOpen Then .PortOpen = False
    End With
Errors:

End Function


Public Function Label_String(ByVal Location As Variant, ByVal WorkArea As Variant, ByVal AccDt As Variant, _
                             ByVal AccSeq As Variant, ByVal SpcNo As Variant, _
                             ByVal Ptid As Variant, ByVal PtNm As Variant, ByVal SpcNm As Variant, _
                             ByVal StoreCd As Variant, ByVal StatFg As Variant, ByVal WardId As Variant, _
                             ByVal OrdDt As Variant, ByVal ColTm As Variant, _
                             ByVal TestNames As Variant, ByVal copycount As Variant, _
                             Optional ByVal AccFg As Boolean = False, Optional ByVal FrozenFg As Boolean = False)

    
    Dim i           As enBarData
    Dim strElemNo   As String
    Dim strTestNm1  As String
    Dim strTestNm2  As String
    Dim strFrozenFg As String
    
    If copycount = 0 Then copycount = 1
    
    If Len(TestNames) > 0 Then
        If Mid(TestNames, Len(TestNames), 1) = "," Then _
            TestNames = Mid(TestNames, 1, Len(TestNames) - 1)
    End If
    
    AccSeq = Trim(AccSeq)
    
    If (mvarAccCheck = "1") And AccFg Then AccSeq = AccSeq & " V"
    If FrozenFg Then strFrozenFg = CS_FrozenFg
    
    SpcNo = AddCheckDigit(CStr(SpcNo))  'check digit 추가
    
'    MsgBox SpcNo & "          체크필드 포함된거"
    
    strTestNm1 = Mid(TestNames, 1, 36) & " "
    strTestNm2 = Mid(TestNames, 37, 36) & " "
        
    Label_String = ""
    Label_String = Label_String & "\1B@z" & vbCrLf
    Label_String = Label_String & "\1B@f" & CS_FotmatNo & vbCrLf
    Label_String = Label_String & "\1Ba" & CS_FotmatNo & mvarLabelLength & mvarLabelTotLength & vbCrLf
    Label_String = Label_String & "\1Bf" & CS_FotmatNo & vbCrLf
      
    'Barcode Label
    If mudtBarcode.PrtFg = "1" Then
'    If mudtBarcode.PrtFg = "1" AND Trim(SpcNo) <> "" Then
        Label_String = Label_String & "\1Bbs" & CS_FotmatNo & "02" & CS_StFg & _
                       mudtBarcode.PosX & mudtBarcode.PosY & mudtBarcode.Length & mudtBarcode.Height & _
                       Format(mudtBarcode.Style, "0#") & NarrowBar & NWRatio & mvarRotation & NotReadable & vbCrLf
    End If
   
    '기타 데이타
    For i = 1 To 14
        With mudtBarData(i)
            If .PrtFg = "1" Then
                strElemNo = Format(i * 2, "0#")
                .ElementNo = strElemNo
                If mudtStatInfo.PrtReverseFg = "1" Then
                    If WardId = mudtStatInfo.ErDeptCd And Location = "응급" And _
                       .Key = mudtStatInfo.ReverseFld Then .ReverseFg = SET_REVERSE
                End If
                Label_String = Label_String & "\1Bds" & CS_FotmatNo & strElemNo & CS_StFg & .PosX & .PosY & _
                               .Length & .FontX & .FontY & Format(mvarRotation, "0#") & .ReverseFg & mvarFontKor & _
                               .BoldFg & vbCrLf
            End If
        End With
    Next
   
    '바코드
    If mudtBarcode.PrtFg = "1" Then
'    If mudtBarcode.PrtFg = "1" AND Trim(SpcNo) <> "" Then
        Label_String = Label_String & "\1Bbw" & CS_FotmatNo & "02" & SpcNo & vbCrLf
    End If
    
    '기타 데이타
    For i = 1 To 14
        With mudtBarData(i)
            If .PrtFg = "1" Then
                Label_String = Label_String & "\1Bdw" & CS_FotmatNo & .ElementNo
                Select Case .Key
                    Case DAT_LOCATION: Label_String = Label_String & Location & vbCrLf
                    Case DAT_WORKAREA: Label_String = Label_String & WorkArea & vbCrLf
                    Case DAT_ACCDT:    Label_String = Label_String & AccDt & vbCrLf
                    Case DAT_ACCSEQ:   Label_String = Label_String & AccSeq & vbCrLf
                    Case DAT_SPCNO:    Label_String = Label_String & SpcNo & vbCrLf
                    Case DAT_DEPT:     Label_String = Label_String & WardId & vbCrLf
                    Case DAT_ORDDT:    Label_String = Label_String & OrdDt & vbCrLf
                    Case DAT_COLTM:    Label_String = Label_String & ColTm & vbCrLf
                    Case DAT_PTNM:     Label_String = Label_String & PtNm & vbCrLf
                    Case DAT_PTID:     Label_String = Label_String & Ptid & vbCrLf
                    Case DAT_SPCNM:    Label_String = Label_String & SpcNm & vbCrLf
                    Case DAT_ORDNM1:   Label_String = Label_String & strTestNm1 & vbCrLf
                    Case DAT_ORDNM2:   Label_String = Label_String & strTestNm2 & vbCrLf
                    Case DAT_FROZEN:   Label_String = Label_String & strFrozenFg & vbCrLf
                End Select
            End If
        End With
    Next
   
    '응급여부
    If (mudtStatInfo.PrtLineFg = "1") And (StatFg = "1") Then
        Label_String = Label_String & "\1Bls" & CS_FotmatNo & "02" & mudtStatInfo.PosX & mudtStatInfo.PosY & _
                       mudtStatInfo.Length & "0000" & mudtStatInfo.Width & vbCrLf        '응급
    End If
    
    Label_String = Label_String & "\1Bq" & Format(copycount, "0###") & vbCrLf
   
End Function

Private Function SQL_BARKIND(ByVal PrjNm) As String
    Dim Rs          As Recordset
    Dim strBarKind  As String
    Dim sSql        As String
    
    If Dir(INIPath) <> "" Then
        strBarKind = medGetINI(PrjNm, "KIND", INIPath)
        SQL_BARKIND = strBarKind
        If SQL_BARKIND <> "" Then Exit Function
    End If
    
    sSql = " SELECT distinct cdval1,field1 from " & objTables.T_COM002 & _
           " WHERE " & DBW("cdindex =", COM2_BAR_CONFIG) & _
           " AND   cdval1 like '" & PrjNm & "%" & "'" & _
           " AND   cdval2  = " & DBS(BAR_KIND)
    
    Set Rs = New Recordset
    
    Rs.Open sSql, DBConn
    If Not Rs.EOF Then
        SQL_BARKIND = Rs.Fields("field1").Value & ""
    End If
    Set Rs = Nothing
End Function

Public Function CitizenPrint(ByVal pLocation As String, ByVal pWorkArea As String, ByVal pAccDt As String, _
                             ByVal pAccSeq As String, ByVal pSpcNo As String, ByVal pPtid As String, _
                             ByVal pPtNm As String, ByVal pSpcNm As String, ByVal pStoreCd As String, _
                             ByVal pStatFg As String, ByVal pWardId As String, ByVal pOrdDt As String, _
                             ByVal pColTm As String, ByVal pTestNames As String, ByVal pCopyCnt As String, _
                             Optional ByVal pAccFg As Boolean = False, Optional ByVal pFrozenFg As Boolean = False)
    Dim strTestNm1  As String
    Dim strTestNm2  As String
    
    Dim CR$, ESC$, K$, KD24$, tmleft$, STX$, SOH$
    Dim tmlen$, tmright$, TMTOT$, TMLEFT1$, SS$
    Dim strSizeX    As String
    Dim strSizeY    As String
    Dim strFrozenFg As String
    Dim ii          As Integer
    Dim strDiv As String
    Dim strTmpSpcno As String
    Dim strBarFor As String
    
    '## 5.0.4: 이상대(2005-02-15)
    '   - Ctizen T4 바코드장비 추가
    '   - QC라벨 기준으로 출력됨, 일반검체에 적용시에는 프로그램 수정해야 함!
On Error GoTo Errors
    With frmComm
        If Not .MyComm.PortOpen Then
            .MyComm.CommPort = mvarPortNo
            .MyComm.Settings = "9600,N,8,1"
            .MyComm.InputLen = 0
            .MyComm.Handshaking = comXOnXoff
            .MyComm.InputMode = comInputModeText
        End If
        If Not .MyComm.PortOpen Then .MyComm.PortOpen = True
    End With
    
    strDiv = "L"
    If Mid(pSpcNo, 1, 1) <> "1" Then strDiv = "B"
    
    If (mvarAccCheck = "1") And pAccFg Then pAccSeq = pAccSeq & " V"   ' 접수여부체크
    
    If pFrozenFg Then strFrozenFg = CS_FrozenFg
    
    If Len(pTestNames) > 0 Then
        If Mid(pTestNames, Len(pTestNames), 1) = "," Then pTestNames = Mid(pTestNames, 1, Len(pTestNames) - 1)
    End If
    
    
    If Len(pTestNames) > 20 Then
        If Mid(Mid(pTestNames, 1, 20), 20) = "," Then strTestNm1 = Mid(pTestNames, 1, 19)
        
        strTestNm1 = Mid(pTestNames, 1, 20) & "..." '30
    Else
        strTestNm1 = Mid(pTestNames, 1, 20)  '30
    End If
    
    SOH = Chr(1)
    STX = Chr(2)
    CR = Chr(13)
    ESC = Chr(27)
        
    With frmComm.MyComm
        .Output = STX & "LD11" & CR
        .Output = "m" & CR
        .Output = "P4S8H14C0000R0000"
        
        'Barcode
        If mudtBarcode.PrtFg = "1" Then
            .Output = "1" & "a" & "42" & "110" & mudtBarcode.PosY & mudtBarcode.PosX & pSpcNo & CR
        End If
        
        '기타 데이타
        For ii = 1 To 14
            If mudtBarData(ii).PrtFg = "1" Then
                If mudtBarData(ii).FontX = FONT_DF Then
                    strSizeX = "1"
                Else
                    strSizeX = mudtBarData(ii).FontX
                End If
                
                If mudtBarData(ii).FontY = FONT_DF Then
                    strSizeY = "1"
                Else
                    strSizeY = mudtBarData(ii).FontY
                End If
                
                Select Case mudtBarData(ii).Key
                    Case DAT_LOCATION:
                            .Output = "1" & ESC$ & strSizeX & strSizeY & "000" & mudtBarData(ii).PosY & mudtBarData(ii).PosX & "KD24" & pLocation & CR    'Location
                    Case DAT_WORKAREA:
                            .Output = "1" & "3" & strSizeX & strSizeY & "000" & mudtBarData(ii).PosY & mudtBarData(ii).PosX & pWorkArea & CR       'WorkArea
                    Case DAT_ACCDT:
                            .Output = "1" & "3" & strSizeX & strSizeY & "000" & mudtBarData(ii).PosY & mudtBarData(ii).PosX & pAccDt & CR          'AccDt
                    Case DAT_ACCSEQ:
                            .Output = "1" & "4" & strSizeX & strSizeY & "000" & mudtBarData(ii).PosY & mudtBarData(ii).PosX & pAccSeq & CR         'AccSeq
                    Case DAT_SPCNO:
                            .Output = "1" & "2" & strSizeX & strSizeY & "000" & mudtBarData(ii).PosY & mudtBarData(ii).PosX & Mid(pSpcNo, 1, 2) & "-" & Format(Mid(pSpcNo, 3), "#") & CR          'SpcNo
                    Case DAT_DEPT:
                            .Output = "1" & "2" & strSizeX & strSizeY & "000" & mudtBarData(ii).PosY & mudtBarData(ii).PosX & pWardId & CR         'WardId
                    Case DAT_ORDDT:
                            .Output = "1" & "2" & strSizeX & strSizeY & "000" & mudtBarData(ii).PosY & mudtBarData(ii).PosX & pOrdDt & CR          'OrdDt
                    Case DAT_COLTM:
                            .Output = "1" & "2" & strSizeX & strSizeY & "000" & mudtBarData(ii).PosY & mudtBarData(ii).PosX & pColTm & CR          'ColTm
                    Case DAT_PTNM:
                            .Output = "1" & "2" & strSizeX & strSizeY & "000" & mudtBarData(ii).PosY & mudtBarData(ii).PosX & strTestNm1 & CR      'TestNm
                    Case DAT_PTID:
                            .Output = "1" & "2" & strSizeX & strSizeY & "000" & mudtBarData(ii).PosY & mudtBarData(ii).PosX & pPtid & CR           'Control ID
                    Case DAT_SPCNM:
                            .Output = "1" & "2" & strSizeX & strSizeY & "000" & mudtBarData(ii).PosY & mudtBarData(ii).PosX & pSpcNm & CR          'Level
                    Case DAT_ORDNM1:
                            .Output = "1" & ESC$ & strSizeX & strSizeY & "000" & mudtBarData(ii).PosY & mudtBarData(ii).PosX & "KD24" & pPtNm & CR           'Control 명
                End Select
            End If
        Next
        DoEvents
        
        .Output = "Q0001" & CR
        .Output = "E"
        If .PortOpen Then .PortOpen = False
    End With
Errors:

End Function

