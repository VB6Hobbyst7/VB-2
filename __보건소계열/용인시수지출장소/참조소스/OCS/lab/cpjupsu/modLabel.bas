Attribute VB_Name = "modLabel"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long


Public GLabelPtno       As String
Public GLabelJeobsuDt   As String

Public GLabelJDt        As String  'General �� �����ð��� ������ �ɶ����� Query������ Ʋ����
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
' C.ITOH s4 ���ڵ� ������ ��� ���� ���� ���α׷�
'
'******************************************************************************
Dim StrData         As String
Dim i               As Integer
'���ڿ� ����'
STX = Chr$(2)
    ESC = Chr$(27)
    CR = Chr$(13)
   
    sBar(4) = convLabnoToComp(Left(sBar(4), 8)) & Right(sBar(4), 7)
    
    Set MSCOM = MFRM.MSComm1
    
    If MSCOM.PortOpen = False Then MSCOM.PortOpen = True
    
    With MSCOM
        .Settings = "9600,N,8,1"
       .Output = STX & "m" & CR
       '.Output = STX & "f400" & CR          ' ' ���� ���� ��ġ ���� ( *0.1mm )  Tear ON ���� ������ �ʿ� ����
       .Output = STX & "L" & CR
       .Output = STX & "m" & CR

       .Output = "D11" & CR         ' ����,���� �� �ȼ�ũ�� ����
       .Output = "H20" & CR         ' �μ� �е� (Printing density, Heat factor) ����
       .Output = "P5" & CR          ' �μ� �ӵ� (Printing speed) ����
       .Output = "S8" & CR          ' ���� �ǵ� �ӵ� (Paper feed speed) ����
        '---------------------------------
        ' ���� ����Ÿ �μ� ��ƾ
        '---------------------------------
        ' Form: RfxyFFFyyyyxxxxTT~T <CR>
        '---------------------------------
'/**************************����Ʈ ����Ʈ***********************************************************************************
'/***       BcDir:      ' �μ���� ( 1=0', 2=90', 3=180', 4=270' )
'/***       BcStyle     ' ���ڵ� ��Ÿ�� (Code3of9, ���ں���=�빮��)
'/***       BcWidth     ' ���� Ȯ�� ���� (�ּ�:1)
'/***       BcHeight    ' ���� Ȯ�� ���� (�ּ�:1)
'/***       BcVHeight   ' ���� ���� ( * 0.1mm )
'/***       BcTop       ' ���� ��ġ (Bottom=Low value(0), Top=High value)
'/***       BcLeft      ' ���� ��ġ (Left=Low value(0), Right=High value)
'/***       BcData      ' ���ڵ� ���� (Data value)
'/********************************************************************************************************************
        
        If CommStrOut("1", "9", "1", "1", "003", "0310", "0015", sBar(0)) = False Then BarCodePrtErr
        If CommStrOut("1", "9", "1", "1", "003", "0310", "0080", sBar(1)) = False Then BarCodePrtErr
        If CommStrOut("1", "9", "1", "1", "003", "0310", "0300", sBar(2)) = False Then BarCodePrtErr
        If CommStrOut("1", "9", "1", "1", "003", "0270", "0080", sBar(3)) = False Then BarCodePrtErr
        If CommBarCodeOut("1", "A", "4", "1", "140", "0100", "0020", sBar(4)) = False Then BarCodePrtErr
        If CommStrOut("1", "9", "1", "1", "003", "0060", "0010", sBar(5)) = False Then BarCodePrtErr
        If CommStrOut("1", ESC, "1", "1", "xxx", "0064", "0180", sBar(6), "K") = False Then BarCodePrtErr  '�ѱ���Ʈ�� �۾Ƽ� �ణ �ø�
        If CommStrOut("1", "9", "1", "1", "003", "0015", "0010", sBar(7)) = False Then BarCodePrtErr
        
'/*****************************�Ϲݹ��� ����Ʈ*******************************************************************************
'/***       BcDir:      ' �μ���� ( 1=0', 2=90', 3=180', 4=270' )
'/***       BcStyle     ' f    ��Ʈ ��Ÿ�� ( 9=Smooth font )
'/***       BcWidth     ' ���� Ȯ�� ���� (�ּ�:1)
'/***       BcHeight    ' ���� Ȯ�� ���� (�ּ�:1)
'/***       BcFont      ' FFF  ��Ʈ ���� ( 3=10pt )
'/***       "KR24";                ' KR24 �ѱ� ���� �μ�ÿ��� �߰��Ǿ�� �� (EPROM ���� ��Ʈ ����)
'/***       BcTop       ' ���� ��ġ (Bottom=Low value(0), Top=High value)
'/***       BcLeft      ' ���� ��ġ (Left=Low value(0), Right=High value)
'/***       BcData      ' ���ڵ� ���� (Data value)
'/********************************************************************************************************************

       .Output = CR                    ' ��� ��
        '---------------------------------
       .Output = "Q0001"                ' �μ� �ż� ����
       .Output = CR                     ' ��� ��
          '*****************************************************************
          ' �ý��� ���� ����
          '*****************************************************************
         .Output = "E"
    End With
    
If MSCOM.PortOpen = True Then MSCOM.PortOpen = False

End Sub


Private Function CommStrOut(BcDir As String, BcStyle As String, BcWidth As String, BcHeight As String, _
                BcFont As String, BcTop As String, BcLeft As String, BcData As String, Optional KrCode As String) As Boolean
                         
'/********************************************************************************************************************
'/***       BcDir:      ' �μ���� ( 1=0', 2=90', 3=180', 4=270' )
'/***       BcStyle     ' f    ��Ʈ ��Ÿ�� ( 9=Smooth font )
'/***       BcWidth     ' ���� Ȯ�� ���� (�ּ�:1)
'/***       BcHeight    ' ���� Ȯ�� ���� (�ּ�:1)
'/***       BcFont      ' FFF  ��Ʈ ���� ( 3=10pt )
'/***       BcTop       ' ���� ��ġ (Bottom=Low value(0), Top=High value)
'/***       BcLeft      ' ���� ��ġ (Left=Low value(0), Right=High value)
'/***       BcData      ' ���ڵ� ���� (Data value)
'/********************************************************************************************************************
                         
        On Error GoTo MscommErr_Rtn
        
        If MSCOM.PortOpen = False Then Exit Function
        If UCase(KrCode) = "K" Then KrCode = "KR24"         '�ѱ��ڵ� ��Ʈ ����
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
'/***       BcDir:      ' �μ���� ( 1=0', 2=90', 3=180', 4=270' )
'/***       BcStyle     ' ���ڵ� ��Ÿ�� (Code3of9, ���ں���=�빮��)
'/***       BcWidth     ' ���� Ȯ�� ���� (�ּ�:1)
'/***       BcHeight    ' ���� Ȯ�� ���� (�ּ�:1)
'/***       BcVHeight   ' ���� ���� ( * 0.1mm )
'/***       BcTop       ' ���� ��ġ (Bottom=Low value(0), Top=High value)
'/***       BcLeft      ' ���� ��ġ (Left=Low value(0), Right=High value)
'/***       BcData      ' ���ڵ� ���� (Data value)
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
   '/���ӵ� ComPort �� Select �ϱ� ���� Function
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
    'Label�����԰� =(����5Cm X ����2Cm)
    '�Է°���Byte = �ѱ�16��(32byte), ����.���� 30(30byte)
    'c04 = Font ��������, k8 = FontSize, o140,220=��ġSet,)
    
    '/---------------------------------------------------------------------------------------------
    'iPortno = GET_ComPort(sComObj)         '���ӵ� ComPort �� ã���ϴ�!."
    'If iPortno = 0 Then
    '    MsgBox "ComPort ������ Ȯ���Ͻʽÿ�!........."
    '    Exit Function
    'End If
    'sComObj.CommPort = iPortno             '���ӵ� ComPort �� ã�Ƽ� Setting
    '/---------------------------------------------------------------------------------------------
    
    If sComObj.PortOpen = True Then sComObj.PortOpen = False
    sComObj.PortOpen = True
    
    sHead = ""        'Init Routine
    sValue1 = ""      'PrintPoint , Font, Size ��.. ����
    sValue2 = ""      'PrintData Vinding
    
    'BarCode Print Initialize
    sHead = sHead & Chr$(2) & Chr$(15) & "T1" & Chr$(3)
    sHead = sHead & Chr$(2) & Chr$(15) & "d10" & Chr$(3)
    sHead = sHead & Chr$(2) & Chr$(15) & "D75" & Chr$(3)           'Cut Line ����
    sHead = sHead & Chr$(2) & Chr$(15) & "F00" & Chr$(3) & Chr$(13)

    'BarCode PrintPoint, Font����,Size���� ����..
    sValue1 = sValue1 & Chr$(2) & Chr$(27) & "C" & Chr$(3)
    sValue1 = sValue1 & Chr$(2) & Chr$(27) & "P" & Chr$(3)
    sValue1 = sValue1 & Chr$(2) & "E3;F3" & Chr$(3)
    sValue1 = sValue1 & Chr$(2) & "H0;o1,1;d3, ;" & Chr$(3)
    
    
    sValue1 = sValue1 & Chr$(2) & "H1;o215,230;f3;c03;h1;w1;d0,50;" & Chr$(3)    'SLipNO1
    sValue1 = sValue1 & Chr$(2) & "H2;o215,270;f3;c03;h1;w1;d0,50;" & Chr$(3)    '��ü
    sValue1 = sValue1 & Chr$(2) & "H3;o215,490;f3;c03;h1;w1;d0,50;" & Chr$(3)    'Emergency/��)
    sValue1 = sValue1 & Chr$(2) & "H4;o185,270;f3;c03;h1;w1;d0,50;" & Chr$(3)    '��ü��ȣ(yyyyyMMdd-no1-no2)
    
    
    'w= BarCode Width, h=Barcode Height
    'c0=39, c2=25, c6=128
    'sValue1 = sValue1 & Chr$(2) & "B5;o147,230;r0;c6,0;f3;i2;h90;w1;d0,50;" & Chr$(3)      'c6 = BarCode Type
    sValue1 = sValue1 & Chr$(2) & "B5;o148,230;r0;c2,0;f3;i2;h90;w2;d0,50;" & Chr$(3)      'c6 = BarCode Type
    
    sValue1 = sValue1 & Chr$(2) & "H6;o45,220;f3;c03;h1;w1;d0,50;" & Chr$(3)    '��Ϲ�ȣ, �̸�, ��
    sValue1 = sValue1 & Chr$(2) & "H7;o23,220;f3;c03;h1;w1;d0,50;" & Chr$(3)    'Item BarText
    'sValue1 = sValue1 & Chr$(2) & "H8;o28,220;f3;c03;h1;w1;d0,50;k9;" & Chr$(3)
    sValue1 = sValue1 & Chr$(2) & "R" & Chr$(3)
    
    '�������ڸ� BarCode�� Length �� ���̱� ���Ͽ� �Լ��� ��ü
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

    '���� �ȳ����淡......
    Call Sleep(2000)      '2������ �������� ���ö��....
                          '�̰͵� �ȵǸ� 3�������� �ѹ� �غ�����.

End Function

Public Function Bar7421_Printing_Sub(ByRef sBar() As String, ByVal nCount As Integer, ByVal sComObj As Object) As Integer
    Dim sHead             As String
    Dim sValue1           As String
    Dim sValue2           As String
    Dim iPortno           As String



    'BarCodePrinter Model = Intermec7421(sammi)    -->  BarCode Printer Model Name(ȸ���̸�)

    'Label�����԰� =(����5Cm X ����2Cm)
    '�Է°���Byte = �ѱ�16��(32byte), ����.���� 30(30byte)
    'c04 = Font ��������, k8 = FontSize, o140,220=��ġSet,)

    '/---------------------------------------------------------------------------------------------
    'iPortno = GET_ComPort(sComObj)         '���ӵ� ComPort �� ã���ϴ�!."
    'If iPortno = 0 Then
    '    MsgBox "ComPort ������ Ȯ���Ͻʽÿ�!........."
    '    Exit Function
    'End If
    'sComObj.CommPort = iPortno             '���ӵ� ComPort �� ã�Ƽ� Setting
    '/---------------------------------------------------------------------------------------------

    If sComObj.PortOpen = True Then sComObj.PortOpen = False
    sComObj.PortOpen = True

    sHead = ""        'Init Routine
    sValue1 = ""      'PrintPoint , Font, Size ��.. ����
    sValue2 = ""      'PrintData Vinding

    'BarCode Print Initialize
    sHead = sHead & Chr$(2) & Chr$(15) & "T1" & Chr$(3)
    sHead = sHead & Chr$(2) & Chr$(15) & "d10" & Chr$(3)
    sHead = sHead & Chr$(2) & Chr$(15) & "D75" & Chr$(3)           'Cut Line ����
    sHead = sHead & Chr$(2) & Chr$(15) & "F00" & Chr$(3) & Chr$(13)

    'BarCode PrintPoint, Font����,Size���� ����..
    sValue1 = sValue1 & Chr$(2) & Chr$(27) & "C" & Chr$(3)
    sValue1 = sValue1 & Chr$(2) & Chr$(27) & "P" & Chr$(3)
    sValue1 = sValue1 & Chr$(2) & "E3;F3" & Chr$(3)
    sValue1 = sValue1 & Chr$(2) & "H0;o1,1;d3, ;" & Chr$(3)


    sValue1 = sValue1 & Chr$(2) & "H1;o215,230;f3;c03;h1;w1;d0,50;" & Chr$(3)    'SLipNO1
    sValue1 = sValue1 & Chr$(2) & "H2;o215,270;f3;c03;h1;w1;d0,50;" & Chr$(3)    '��ü
    sValue1 = sValue1 & Chr$(2) & "H3;o215,490;f3;c03;h1;w1;d0,50;" & Chr$(3)    'Emergency/��)
    sValue1 = sValue1 & Chr$(2) & "H4;o185,270;f3;c03;h1;w1;d0,50;" & Chr$(3)    '��ü��ȣ(yyyyyMMdd-no1-no2)


    'w= BarCode Width, h=Barcode Height
    'c0=39, c2=25, c6=128
    'sValue1 = sValue1 & Chr$(2) & "B5;o147,230;r0;c6,0;f3;i2;h90;w1;d0,50;" & Chr$(3)      'c6 = BarCode Type
    sValue1 = sValue1 & Chr$(2) & "B5;o148,230;r0;c2,0;f3;i2;h90;w2;d0,50;" & Chr$(3)      'c6 = BarCode Type

    sValue1 = sValue1 & Chr$(2) & "H6;o45,220;f3;c03;h1;w1;d0,50;" & Chr$(3)    '��Ϲ�ȣ, �̸�, ��
    sValue1 = sValue1 & Chr$(2) & "H7;o23,220;f3;c03;h1;w1;d0,50;" & Chr$(3)    'Item BarText
    'sValue1 = sValue1 & Chr$(2) & "H8;o28,220;f3;c03;h1;w1;d0,50;k9;" & Chr$(3)
    sValue1 = sValue1 & Chr$(2) & "R" & Chr$(3)

    '�������ڸ� BarCode�� Length �� ���̱� ���Ͽ� �Լ��� ��ü
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

    '���� �ȳ����淡......
    Call Sleep(2000)      '2������ �������� ���ö��....
                          '�̰͵� �ȵǸ� 3�������� �ѹ� �غ�����.

End Function

Public Function convSLipYageo(ByVal sSLipno1 As String) As String
        
    'SLipno1 �� �ӻ󺴸������� ���ϴ� ������.............
    'SLipno1 �� Conversion ��.
    
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
