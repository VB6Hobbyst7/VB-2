VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   405
      Left            =   840
      TabIndex        =   1
      Top             =   210
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   1755
      Left            =   540
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   3405
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�������� ������ ������ ������ �Ӹ��ƴ϶� SOAP Envelope �� �����Ͽ� ����� �� �ִ� ��� ������ ���� ���� �ؾ� �Ѵ�.
Private Const ENC = "http://schemas.xmlsoap.org/soap/encoding/"
Private Const XSI = "http://www.w3.org/1999/XMLSchema-instance"
Private Const XSD = "http://www.w3.org/1999/XMLSchema"

Private Sub Command1_Click()
''�������� ������ ������ ������ �Ӹ��ƴ϶� SOAP Envelope �� �����Ͽ� ����� �� �ִ� ��� ������ ���� ���� �ؾ� �Ѵ�.
'        Private Const ENC = "http://schemas.xmlsoap.org/soap/encoding/"
'        Private Const XSI = "http://www.w3.org/1999/XMLSchema-instance"
'        Private Const XSD = "http://www.w3.org/1999/XMLSchema"
        
'        url = "http://www.soapclient.com/interop/InteropB.wsdl"
        url = "http://10.47.14.52:8009/HL7IFWebService/WebService.asmx?wsdl"
        URI = "urn:QuotationService"
        Method = "getQuotationsByAuthor"

'    defSvc = "SQLDataInterop"
'    defPrt = "InteropTestPort"
'
'�������� SOAP Connector, Serializer and Reader �ʱ�ȭ �ؾ� �Ѵ�. �� Ŀ���ʹ� HTTP Ŀ�ؼ����� �ٷ� ���̴�. Serializer �� �������� SOAP Envelope�� ���� �ϴ°��� ���� �� ���̸� Reader�� �������� �� ����� ����� �� �ְ� �����ش�.
        Dim Connector As SoapConnector30
        Dim Serializer As SoapSerializer30
        Dim Reader As SoapReader30
        
        Set Connector = New HttpConnector30
        Set Serializer = New SoapSerializer30
        Set Reader = New SoapReader30

'������ SOAP Server�� Ŀ���͸� �غ� �ؾ� �Ѵ�. "SoapAction" ����Ÿ�� �������̵忡���� �翬���� ������, �� ������ �� ��� ��� ���� �� �� �ִ�. �̰��� �������� SOAP �޽����� �а� ������� �� �� �� ���� Ȯ���� �� �ִ� URI�� �޼ҵ� �̸��� ��Ʈ �ϴ� ���� �����̴�.
        Connector.Property("EndPointURL") = url
        Call Connector.Connect
        Connector.Property("SoapAction") = URI & "#" & Method
        Call Connector.BeginMessage

'�������� Serialize�� Connector�� ���ý�Ű��.
        Serializer.Init Connector.InputStream

'SOAP Envelope�� �����ϰ� Encoding �� XML -Schema�� ����Ѵ�.
        Serializer.StartEnvelope , ENC
        Serializer.SoapNamespace "xsi", XSI
        Serializer.SoapNamespace "SOAP-ENC", ENC
        Serializer.SoapNamespace "xsd", XSD

'�޽����� �ٵ� �����Ѵ�. ��Ʈ ���Ҵ� �׻� ���� URI�� �޼ҵ忩�� �Ѵ�.
        Serializer.StartBody
        Serializer.StartElement Method, URI, , "method"

'��Ʈ ���ҿ� ���ϵ�� ������ �޼ҵ� �Ķ���Ϳ� ���� ����� �Ѵ�.
        Serializer.StartElement "Author"
        Serializer.SoapAttribute "type", , "xsd:string", "xsi"
        Serializer.WriteString "Wilde, Oscar"
        Serializer.EndElement

'��Ʈ ������ ���� body�� envelope�� ��ġ�Ѵ�.
        Serializer.EndElement
        Serializer.EndBody
        Serializer.EndEnvelope

'���´ٸ� �޽��� ���� ���� �ؾ� �Ѵ�.
        Connector.EndMessage

'�����ȿ� ����� �ε� �϶�.
        Reader.Load Connector.OutputStream

        If Not Reader.Fault Is Nothing Then
          MsgBox Reader.FaultString.Text, vbExclamation
        Else
          'Set Result = Reader.Dom
          '//parse the DOM to extract the result set
        End If


End Sub

Private Sub Form_Load()
'Exit Sub
'Dim SoapClient As SoapClient30
'Dim defWSDL, defURL, defSvc, defPrt As String
'
'Set SoapClient = New SoapClient30
'
'  defWSDL = "D:\������Ʈ\���������Ǽ�\SOAP\��������\�������� Sample\WebService.asmx"
''  defURL = "http://10.47.14.52:8009/HL7IFWebService/WebService.asmx"
''  defWSDL = "http://10.47.14.52:8009/HL7IFWebService/WebService.asmx?wsdl"
'
''  defWSDL = "http://192.168.123.4:1029/vbService/Service.asmx"
'  defURL = "http://192.168.123.4:1029/vbService/Service.asmx"
'  defSvc = "WebService"
'  defPrt = "WebServiceSoap"
'
'    defWSDL = "http://www.soapclient.com/interop/InteropB.wsdl"
'    defURL = "http://192.168.123.4:1029/vbService/Service.asmx"
'    defSvc = "SQLDataInterop"
'    defPrt = "InteropTestPort"
'
'    SoapClient.ClientProperty("ServerHTTPRequest") = True
'
'    Call SoapClient.MSSoapInit(defWSDL, defSvc, defPrt)
'
'
'
''wscript.echo SoapClient.AddNumbers(2, 3) '�����񽺿� ���ǵ� �޼ҵ� ȣ��
'
''Print low - Level; interface; ����
'Dim Serializer As SoapSerializer30 '������ �����͸� SOAP XML���·�
'Dim Reader As SoapReader30 '���� �����͸� XML ���·�
'
''Set Connector = New HttpConnector30 '�ش� �ּҷ� ����
'
''Connector.Property("EndPointURL") = "http://www.soapclient.com/interop/InteropB.wsdl" '"http://www.xxxx.com/webservice.php"
''Connector.Connect
''Connector.Property("SoapAction") = "uri:" & Method
''Connector.BeginMessage
'
'Set Serializer = New SoapSerializer30
''Serializer.Init Connector.InputStream
'
''MsgBox ("SOAP ��� �����ͻ���")
'
'Serializer.StartEnvelope
'Serializer.StartBody
''Serializer.StartElement "getRecommendation", CALC_NS, , "nstemp"
'Serializer.StartElement "data"
'Serializer.WriteString Text1.Text
'Serializer.EndElement
'Serializer.EndElement
'Serializer.EndBody
'Serializer.EndEnvelope
'Connector.EndMessage
'
''On Error Resume Next
'
''MsgBox ("SOAP ��� ��� ���")
'
'Set Reader = New SoapReader30
'Reader.Load Connector.OutputStream
'Text1.Text = Reader.Body.xml
'MsgBox Reader.Body.xml

End Sub
