Attribute VB_Name = "modLabel"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long


Public GLabelPtno       As String
Public GLabelJeobsuDt   As String

Public GLabelJDt        As String  'General 의 접수시간이 조건이 될때에는 Query문장이 틀려짐
Public GLabelJT1        As String
Public GLabelJT2        As String
Public GLabelLabno1     As Integer
Public GLabelLabno2     As Integer
Public GLabelLoadCheck  As String


Public GVarPoint    As Integer

Public Type LabelBar
    Title(100)       As String
    Ptno(100)        As String
    JeobsuDt(100)    As String
    sLipno1(100)     As String
    Slipno2(100)     As String
    DeptCode(100)    As String
    BarText(100)     As String
    Yg(100)          As String
    SampleCd(100)    As String
    ReporCd(100)     As String
    Er(100)          As String
    ChUnit(100)      As String
End Type

Public LabelString  As LabelBar

Public Type LabelBar1
    Title(100)       As String
    Ptno(100)        As String
    JeobsuDt(100)    As String
    sLipno1(100)     As String
    Slipno2(100)     As String
    DeptCode(100)    As String
    BarText(100)     As String
    Yg(100)          As String
    SampleCd(100)    As String
    ReporCd(100)     As String
    Er(100)          As String
    ChUnit(100)      As String
End Type

Public LabelString1  As LabelBar1

Dim STX                 As String
Dim ESC                 As String
Dim CR                  As String

Public MSCOM            As MSComm
Public Sub BarCodePrint(ByRef sBar() As String, ByVal nCount As Integer, MFRM As Form)
'******************************************************************************
'
' C.ITOH s4 바코드 프린터 제어를 위한 샘플 프로그램
'
'******************************************************************************
Dim StrData         As String
Dim i               As Integer
'문자열 정의'
STX = Chr$(2)
    ESC = Chr$(27)
    CR = Chr$(13)
   
    sBar(4) = convLabnoToComp(Left(sBar(4), 8)) & Right(sBar(4), 7)
    
    Set MSCOM = MFRM.MSComm1
    
    If MSCOM.PortOpen = False Then MSCOM.PortOpen = True
    
    With MSCOM
        .Settings = "9600,N,8,1"
       .Output = STX & "m" & CR
       '.Output = STX & "f400" & CR          ' ' 용지 배출 위치 지정 ( *0.1mm )  Tear ON 모드로 지정시 필요 없당
       .Output = STX & "L" & CR
       .Output = STX & "m" & CR

       .Output = "D11" & CR         ' 가로,세로 의 픽셀크기 지정
       .Output = "H20" & CR         ' 인쇄 밀도 (Printing density, Heat factor) 지정
       .Output = "P5" & CR          ' 인쇄 속도 (Printing speed) 지정
       .Output = "S8" & CR          ' 용지 피드 속도 (Paper feed speed) 지정
        '---------------------------------
        ' 문자 데이타 인쇄 루틴
        '---------------------------------
        ' Form: RfxyFFFyyyyxxxxTT~T <CR>
        '---------------------------------
'/**************************바코트 프린트***********************************************************************************
'/***       BcDir:      ' 인쇄방향 ( 1=0', 2=90', 3=180', 4=270' )
'/***       BcStyle     ' 바코드 스타일 (Code3of9, 숫자보임=대문자)
'/***       BcWidth     ' 가로 확대 지정 (최소:1)
'/***       BcHeight    ' 세로 확대 지정 (최소:1)
'/***       BcVHeight   ' 세로 높이 ( * 0.1mm )
'/***       BcTop       ' 세로 위치 (Bottom=Low value(0), Top=High value)
'/***       BcLeft      ' 가로 위치 (Left=Low value(0), Right=High value)
'/***       BcData      ' 바코드 내용 (Data value)
'/********************************************************************************************************************
        
        If CommStrOut("1", "9", "1", "1", "003", "0310", "0015", sBar(0)) = False Then BarCodePrtErr
        If CommStrOut("1", "9", "1", "1", "003", "0310", "0080", sBar(1)) = False Then BarCodePrtErr
        If CommStrOut("1", "9", "1", "1", "003", "0310", "0300", sBar(2)) = False Then BarCodePrtErr
        If CommStrOut("1", "9", "1", "1", "003", "0270", "0080", sBar(3)) = False Then BarCodePrtErr
        If CommBarCodeOut("1", "A", "4", "1", "140", "0100", "0020", sBar(4)) = False Then BarCodePrtErr
        If CommStrOut("1", "9", "1", "1", "003", "0060", "0010", sBar(5)) = False Then BarCodePrtErr
        If CommStrOut("1", ESC, "1", "1", "xxx", "0064", "0180", sBar(6), "K") = False Then BarCodePrtErr  '한글폰트가 작아서 약간 올림
        If CommStrOut("1", "9", "1", "1", "003", "0015", "0010", sBar(7)) = False Then BarCodePrtErr
        
'/*****************************일반문자 프린트*******************************************************************************
'/***       BcDir:      ' 인쇄방향 ( 1=0', 2=90', 3=180', 4=270' )
'/***       BcStyle     ' f    폰트 스타일 ( 9=Smooth font )
'/***       BcWidth     ' 가로 확대 지정 (최소:1)
'/***       BcHeight    ' 세로 확대 지정 (최소:1)
'/***       BcFont      ' FFF  폰트 지정 ( 3=10pt )
'/***       "KR24";                ' KR24 한글 문자 인쇄시에는 추가되어야 함 (EPROM 에서 폰트 선택)
'/***       BcTop       ' 세로 위치 (Bottom=Low value(0), Top=High value)
'/***       BcLeft      ' 가로 위치 (Left=Low value(0), Right=High value)
'/***       BcData      ' 바코드 내용 (Data value)
'/********************************************************************************************************************

       .Output = CR                    ' 명령 끝
        '---------------------------------
       .Output = "Q0001"                ' 인쇄 매수 지정
       .Output = CR                     ' 명령 끝
          '*****************************************************************
          ' 시스템 모드로 복귀
          '*****************************************************************
         .Output = "E"
    End With
    
If MSCOM.PortOpen = True Then MSCOM.PortOpen = False

End Sub


Private Function CommStrOut(BcDir As String, BcStyle As String, BcWidth As String, BcHeight As String, _
                BcFont As String, BcTop As String, BcLeft As String, BcData As String, Optional KrCode As String) As Boolean
                         
'/********************************************************************************************************************
'/***       BcDir:      ' 인쇄방향 ( 1=0', 2=90', 3=180', 4=270' )
'/***       BcStyle     ' f    폰트 스타일 ( 9=Smooth font )
'/***       BcWidth     ' 가로 확대 지정 (최소:1)
'/***       BcHeight    ' 세로 확대 지정 (최소:1)
'/***       BcFont      ' FFF  폰트 지정 ( 3=10pt )
'/***       BcTop       ' 세로 위치 (Bottom=Low value(0), Top=High value)
'/***       BcLeft      ' 가로 위치 (Left=Low value(0), Right=High value)
'/***       BcData      ' 바코드 내용 (Data value)
'/********************************************************************************************************************
                         
        On Error GoTo MscommErr_Rtn
        
        If MSCOM.PortOpen = False Then Exit Function
        If UCase(KrCode) = "K" Then KrCode = "KR24"         '한글코드 폰트 지정
        MSCOM.Output = BcDir & BcStyle & BcWidth & BcHeight & BcFont & BcTop & BcLeft & KrCode & BcData & CR
        CommStrOut = True
Exit Function

MscommErr_Rtn:
    MsgBox Err.Number & ": " & Err.Description
    CommStrOut = False
End Function
                         
Private Function CommBarCodeOut(BcDir As String, BcStyle As String, BcWidth As String, BcHeight As String, _
                                BcVHeight As String, BcTop As String, BcLeft As String, BcData As String) As Boolean
'/********************************************************************************************************************
'/***       BcDir:      ' 인쇄방향 ( 1=0', 2=90', 3=180', 4=270' )
'/***       BcStyle     ' 바코드 스타일 (Code3of9, 숫자보임=대문자)
'/***       BcWidth     ' 가로 확대 지정 (최소:1)
'/***       BcHeight    ' 세로 확대 지정 (최소:1)
'/***       BcVHeight   ' 세로 높이 ( * 0.1mm )
'/***       BcTop       ' 세로 위치 (Bottom=Low value(0), Top=High value)
'/***       BcLeft      ' 가로 위치 (Left=Low value(0), Right=High value)
'/***       BcData      ' 바코드 내용 (Data value)
'/********************************************************************************************************************

        On Error GoTo MscommErr_Rtn
        
        If MSCOM.PortOpen = False Then Exit Function
        
        MSCOM.Output = BcDir & BcStyle & BcWidth & BcHeight & BcVHeight & BcTop & BcLeft & BcData & CR
        CommBarCodeOut = True

Exit Function

MscommErr_Rtn:
    MsgBox Err.Number & ": " & Err.Description
    CommBarCodeOut = False
    
End Function


Private Sub BarCodePrtErr()
    If MSCOM.PortOpen = True Then MSCOM.PortOpen = False
    End
End Sub

Public Sub LabelStringClear()
    
    For i = 0 To 100
        LabelString.Title(i) = ""
        LabelString.Ptno(i) = ""
        LabelString.JeobsuDt(i) = ""
        LabelString.sLipno1(i) = ""
        LabelString.Slipno2(i) = ""
        LabelString.DeptCode(i) = ""
        LabelString.BarText(i) = ""
        LabelString.Yg(i) = ""
        LabelString.SampleCd(i) = ""
        LabelString.ReporCd(i) = ""
        LabelString.Er(i) = ""
    Next
    
End Sub
Public Sub LabelString1Clear()
    
    For i = 0 To 100
        LabelString1.Title(i) = ""
        LabelString1.Ptno(i) = ""
        LabelString1.JeobsuDt(i) = ""
        LabelString1.sLipno1(i) = ""
        LabelString1.Slipno2(i) = ""
        LabelString1.DeptCode(i) = ""
        LabelString1.BarText(i) = ""
        LabelString1.Yg(i) = ""
        LabelString1.SampleCd(i) = ""
        LabelString1.ReporCd(i) = ""
        LabelString1.Er(i) = ""
    Next

End Sub
Public Function isArrayMaxReturn(ByRef arrayRet() As String) As Integer
    Dim nCnt    As Integer
    
    isArrayMaxReturn = 0
    For nCnt = LBound(arrayRet) To UBound(arrayRet)
        If arrayRet(nCnt) = "" Then
            isArrayMaxReturn = nCnt
            Exit For
        End If
    Next
    
    
End Function
Public Function isArrayText(ByRef arrayReturn() As String, ByVal sText As String) As Integer
    Dim nCnt    As Integer
    
    'False =  0
    'True  = -1
    GVarPoint = 0
    isArrayText = False
    
    For nCnt = LBound(arrayReturn) To UBound(arrayReturn)
        If Trim(arrayReturn(nCnt)) = Trim(sText) Then
            isArrayText = True
            GVarPoint = nCnt
            Exit For
        End If
    Next

End Function

Public Function GET_ComPort(ByVal sComObj As Object) As Integer
   '/접속된 ComPort 를 Select 하기 위한 Function
    Dim iPort       As Integer
    
    GET_ComPort = 0
    On Error GoTo Select_Port
    
    For iPort = 1 To 2
        sComObj.CommPort = iPort
        If sComObj.PortOpen = True Then sComObj.PortOpen = False
    Next
    
    For iPort = 1 To 2
        sComObj.CommPort = iPort
        sComObj.PortOpen = True
        
        If sComObj.CTSHolding = True Then
            GET_ComPort = sComObj.CommPort
        End If
        sComObj.PortOpen = False
    Next
    Exit Function
    
    
Select_Port:
    If Err.Number = 8005 Then
        GET_ComPort = 0
    End If
    Exit Function

End Function

Public Function Bar7421_Printing_Donner_Sub(ByRef sBar() As String, ByVal nCount As Integer, ByVal sComObj As Object) As Integer


    Dim sHead             As String
    Dim sValue1           As String
    Dim sValue2           As String
    Dim iPortno           As String
    
        
        
    'BarCodePrinter Model = Intermec7421(sammi)
    'Label용지규격 =(가로5Cm X 세로2Cm)
    '입력가능Byte = 한글16자(32byte), 영문.숫자 30(30byte)
    'c04 = Font 종류구분, k8 = FontSize, o140,220=위치Set,)
    
    '/---------------------------------------------------------------------------------------------
    'iPortno = GET_ComPort(sComObj)         '접속된 ComPort 를 찾슶니다!."
    'If iPortno = 0 Then
    '    MsgBox "ComPort 연결을 확인하십시오!........."
    '    Exit Function
    'End If
    'sComObj.CommPort = iPortno             '접속된 ComPort 를 찾아서 Setting
    '/---------------------------------------------------------------------------------------------
    
    If sComObj.PortOpen = True Then sComObj.PortOpen = False
    sComObj.PortOpen = True
    
    sHead = ""        'Init Routine
    sValue1 = ""      'PrintPoint , Font, Size 등.. 정의
    sValue2 = ""      'PrintData Vinding
    
    'BarCode Print Initialize
    sHead = sHead & Chr$(2) & Chr$(15) & "T1" & Chr$(3)
    sHead = sHead & Chr$(2) & Chr$(15) & "d10" & Chr$(3)
    sHead = sHead & Chr$(2) & Chr$(15) & "D75" & Chr$(3)           'Cut Line 맞춤
    sHead = sHead & Chr$(2) & Chr$(15) & "F00" & Chr$(3) & Chr$(13)

    'BarCode PrintPoint, Font종류,Size등을 정의..
    sValue1 = sValue1 & Chr$(2) & Chr$(27) & "C" & Chr$(3)
    sValue1 = sValue1 & Chr$(2) & Chr$(27) & "P" & Chr$(3)
    sValue1 = sValue1 & Chr$(2) & "E3;F3" & Chr$(3)
    sValue1 = sValue1 & Chr$(2) & "H0;o1,1;d3, ;" & Chr$(3)
    
    
    sValue1 = sValue1 & Chr$(2) & "H1;o215,230;f3;c03;h1;w1;d0,50;" & Chr$(3)    'SLipNO1
    sValue1 = sValue1 & Chr$(2) & "H2;o215,270;f3;c03;h1;w1;d0,50;" & Chr$(3)    '검체
    sValue1 = sValue1 & Chr$(2) & "H3;o215,490;f3;c03;h1;w1;d0,50;" & Chr$(3)    'Emergency/외)
    sValue1 = sValue1 & Chr$(2) & "H4;o185,270;f3;c03;h1;w1;d0,50;" & Chr$(3)    '검체번호(yyyyyMMdd-no1-no2)
    
    
    'w= BarCode Width, h=Barcode Height
    'c0=39, c2=25, c6=128
    'sValue1 = sValue1 & Chr$(2) & "B5;o147,230;r0;c6,0;f3;i2;h90;w1;d0,50;" & Chr$(3)      'c6 = BarCode Type
    sValue1 = sValue1 & Chr$(2) & "B5;o148,230;r0;c2,0;f3;i2;h90;w2;d0,50;" & Chr$(3)      'c6 = BarCode Type
    
    sValue1 = sValue1 & Chr$(2) & "H6;o45,220;f3;c03;h1;w1;d0,50;" & Chr$(3)    '등록번호, 이름, 과
    sValue1 = sValue1 & Chr$(2) & "H7;o23,220;f3;c03;h1;w1;d0,50;" & Chr$(3)    'Item BarText
    'sValue1 = sValue1 & Chr$(2) & "H8;o28,220;f3;c03;h1;w1;d0,50;k9;" & Chr$(3)
    sValue1 = sValue1 & Chr$(2) & "R" & Chr$(3)
    
    '접수일자를 BarCode의 Length 를 줄이기 위하여 함수로 대체
    'sBar(4) = convLabnoToComp(Left(sBar(4), 8)) & Right(sBar(4), 7)
    sBar(4) = sBar(4)
    
    sValue2 = sValue2 & Chr$(2) & Chr$(27) & "E3" & Chr$(24) & Chr$(3)
    sValue2 = sValue2 & Chr$(2) & sBar(0) & Chr$(13) & Chr$(3)           'H1
    sValue2 = sValue2 & Chr$(2) & sBar(1) & Chr$(13) & Chr$(3)           'H2
    sValue2 = sValue2 & Chr$(2) & sBar(2) & Chr$(13) & Chr$(3)           'H3
    sValue2 = sValue2 & Chr$(2) & sBar(3) & Chr$(13) & Chr$(3)           'H4
    sValue2 = sValue2 & Chr$(2) & sBar(4) & Chr$(13) & Chr$(3)           'H5 (BarCodeData)
    sValue2 = sValue2 & Chr$(2) & sBar(5) & Chr$(13) & Chr$(3)           'H6
    sValue2 = sValue2 & Chr$(2) & sBar(6) & Chr$(13) & Chr$(3)           'H7
    sValue2 = sValue2 & Chr$(2) & Chr$(30) & nCount & Chr$(23) & Chr$(3)


    sComObj.Output = sHead
    
    While sComObj.OutBufferCount > 0
        DoEvents
    Wend
    
    sComObj.Output = sValue1
    While sComObj.OutBufferCount > 0
        DoEvents
    Wend
    
    sComObj.Output = sValue2
    While sComObj.OutBufferCount > 0
        DoEvents
    Wend

    If sComObj.PortOpen = True Then sComObj.PortOpen = False

    '가끔 안나오길래......
    Call Sleep(2000)      '2초정도 쉬엄쉬엄 가시라고....
                          '이것도 안되면 3초정도로 한번 해봐야지.

End Function

Public Function Bar7421_Printing_Sub(ByRef sBar() As String, ByVal nCount As Integer, ByVal sComObj As Object) As Integer
    Dim sHead             As String
    Dim sValue1           As String
    Dim sValue2           As String
    Dim iPortno           As String



    'BarCodePrinter Model = Intermec7421(sammi)    -->  BarCode Printer Model Name(회사이름)

    'Label용지규격 =(가로5Cm X 세로2Cm)
    '입력가능Byte = 한글16자(32byte), 영문.숫자 30(30byte)
    'c04 = Font 종류구분, k8 = FontSize, o140,220=위치Set,)

    '/---------------------------------------------------------------------------------------------
    'iPortno = GET_ComPort(sComObj)         '접속된 ComPort 를 찾슶니다!."
    'If iPortno = 0 Then
    '    MsgBox "ComPort 연결을 확인하십시오!........."
    '    Exit Function
    'End If
    'sComObj.CommPort = iPortno             '접속된 ComPort 를 찾아서 Setting
    '/---------------------------------------------------------------------------------------------

    If sComObj.PortOpen = True Then sComObj.PortOpen = False
    sComObj.PortOpen = True

    sHead = ""        'Init Routine
    sValue1 = ""      'PrintPoint , Font, Size 등.. 정의
    sValue2 = ""      'PrintData Vinding

    'BarCode Print Initialize
    sHead = sHead & Chr$(2) & Chr$(15) & "T1" & Chr$(3)
    sHead = sHead & Chr$(2) & Chr$(15) & "d10" & Chr$(3)
    sHead = sHead & Chr$(2) & Chr$(15) & "D75" & Chr$(3)           'Cut Line 맞춤
    sHead = sHead & Chr$(2) & Chr$(15) & "F00" & Chr$(3) & Chr$(13)

    'BarCode PrintPoint, Font종류,Size등을 정의..
    sValue1 = sValue1 & Chr$(2) & Chr$(27) & "C" & Chr$(3)
    sValue1 = sValue1 & Chr$(2) & Chr$(27) & "P" & Chr$(3)
    sValue1 = sValue1 & Chr$(2) & "E3;F3" & Chr$(3)
    sValue1 = sValue1 & Chr$(2) & "H0;o1,1;d3, ;" & Chr$(3)


    sValue1 = sValue1 & Chr$(2) & "H1;o215,230;f3;c03;h1;w1;d0,50;" & Chr$(3)    'SLipNO1
    sValue1 = sValue1 & Chr$(2) & "H2;o215,270;f3;c03;h1;w1;d0,50;" & Chr$(3)    '검체
    sValue1 = sValue1 & Chr$(2) & "H3;o215,490;f3;c03;h1;w1;d0,50;" & Chr$(3)    'Emergency/외)
    sValue1 = sValue1 & Chr$(2) & "H4;o185,270;f3;c03;h1;w1;d0,50;" & Chr$(3)    '검체번호(yyyyyMMdd-no1-no2)


    'w= BarCode Width, h=Barcode Height
    'c0=39, c2=25, c6=128
    'sValue1 = sValue1 & Chr$(2) & "B5;o147,230;r0;c6,0;f3;i2;h90;w1;d0,50;" & Chr$(3)      'c6 = BarCode Type
    sValue1 = sValue1 & Chr$(2) & "B5;o148,230;r0;c2,0;f3;i2;h90;w2;d0,50;" & Chr$(3)      'c6 = BarCode Type

    sValue1 = sValue1 & Chr$(2) & "H6;o45,220;f3;c03;h1;w1;d0,50;" & Chr$(3)    '등록번호, 이름, 과
    sValue1 = sValue1 & Chr$(2) & "H7;o23,220;f3;c03;h1;w1;d0,50;" & Chr$(3)    'Item BarText
    'sValue1 = sValue1 & Chr$(2) & "H8;o28,220;f3;c03;h1;w1;d0,50;k9;" & Chr$(3)
    sValue1 = sValue1 & Chr$(2) & "R" & Chr$(3)

    '접수일자를 BarCode의 Length 를 줄이기 위하여 함수로 대체
    sBar(4) = convLabnoToComp(Left(sBar(4), 8)) & Right(sBar(4), 7)
    'sBar(4) = sBar(4)

    sValue2 = sValue2 & Chr$(2) & Chr$(27) & "E3" & Chr$(24) & Chr$(3)
    sValue2 = sValue2 & Chr$(2) & sBar(0) & Chr$(13) & Chr$(3)           'H1
    sValue2 = sValue2 & Chr$(2) & sBar(1) & Chr$(13) & Chr$(3)           'H2
    sValue2 = sValue2 & Chr$(2) & sBar(2) & Chr$(13) & Chr$(3)           'H3
    sValue2 = sValue2 & Chr$(2) & sBar(3) & Chr$(13) & Chr$(3)           'H4
    sValue2 = sValue2 & Chr$(2) & sBar(4) & Chr$(13) & Chr$(3)           'H5 (BarCodeData)
    sValue2 = sValue2 & Chr$(2) & sBar(5) & Chr$(13) & Chr$(3)           'H6
    sValue2 = sValue2 & Chr$(2) & sBar(6) & Chr$(13) & Chr$(3)           'H7
    sValue2 = sValue2 & Chr$(2) & Chr$(30) & nCount & Chr$(23) & Chr$(3)


    sComObj.Output = sHead

    While sComObj.OutBufferCount > 0
        DoEvents
    Wend

    sComObj.Output = sValue1
    While sComObj.OutBufferCount > 0
        DoEvents
    Wend

    sComObj.Output = sValue2
    While sComObj.OutBufferCount > 0
        DoEvents
    Wend

    If sComObj.PortOpen = True Then sComObj.PortOpen = False

    '가끔 안나오길래......
    Call Sleep(2000)      '2초정도 쉬엄쉬엄 가시라고....
                          '이것도 안되면 3초정도로 한번 해봐야지.

End Function

Public Function convSLipYageo(ByVal sSLipno1 As String) As String
        
    'SLipno1 을 임상병리과에서 원하는 것으로.............
    'SLipno1 을 Conversion 함.
    
    convSLipYageo = ""
    Select Case sSLipno1
        Case "11": convSLipYageo = "H1"
        Case "12": convSLipYageo = "H2"
        Case "13": convSLipYageo = "U1"
        Case "14": convSLipYageo = "F0"
        Case "15": convSLipYageo = "H3"
        Case "21": convSLipYageo = "C1"
        Case "22": convSLipYageo = "C2"
        Case "23": convSLipYageo = "GA"
        Case "24": convSLipYageo = "U2"
        Case "31": convSLipYageo = "S1"
        Case "32": convSLipYageo = "S2"
        Case "33": convSLipYageo = "S3"
        Case "34": convSLipYageo = "G1"
        Case "35": convSLipYageo = "G2"
        Case "41": convSLipYageo = "M2"
        Case "42": convSLipYageo = "M1"
        Case "43": convSLipYageo = "P0"
        Case "44": convSLipYageo = "M3"
        Case "45": convSLipYageo = "V0"
        Case "51": convSLipYageo = "B1"
        Case Else: convSLipYageo = sSLipno1
    End Select

End Function


Public Function convLabnoToExpand(ByVal sComp5 As String) As String
    
    convLabnoToExpand = Format(DateAdd("d", Val(sComp5), "2000-10-01"), "YYYYMMDD")
        
    
End Function

Public Function convLabnoToComp(ByVal sYear8 As String) As String
    Dim sconvYear      As String
    
    sconvYear = Left(sYear8, 4) & "-" & Mid(sYear8, 5, 2) & "-" & Mid(sYear8, 7)
    
    convLabnoToComp = Format(DateDiff("d", "2000-10-01", sconvYear), "00000")
    
End Function


Public Function GET_COLLDate(ByVal sJeobsuDt As String, ByVal iSLno1 As Integer, ByVal iSLno2 As Integer) As String
    Dim adoColl     As ADODB.Recordset
    
    strSql = ""
    strSql = strSql & " SELECT TO_CHAR(b.CollDate,'yyyy-MM-dd') COLLDate"
    strSql = strSql & " FROM   TWEXAM_General a,"
    strSql = strSql & "        TW_MIS_EXAM.TWEXAM_Order   b"
    strSql = strSql & " WHERE  a.JeobsuDt = TO_DATE('" & sJeobsuDt & "','yyyy-MM-dd')"
    strSql = strSql & " AND    a.SLipno1  = " & iSLno1
    strSql = strSql & " AND    a.SLipno2  = " & iSLno2
    strSql = strSql & " AND    a.JeobsuDt = b.JeobsuDt(+)"
    strSql = strSql & " AND    a.SLipno1  = b.SLipno1(+)"
    strSql = strSql & " AND    a.Orderno  = b.Orderno(+)"
    
    If adoSetOpen(strSql, adoColl) Then
        GET_COLLDate = adoColl.Fields("COLLDate").Value & ""
        Call adoSetClose(adoColl)
    Else
        GET_COLLDate = ""
        Exit Function
    End If
    
    
End Function
