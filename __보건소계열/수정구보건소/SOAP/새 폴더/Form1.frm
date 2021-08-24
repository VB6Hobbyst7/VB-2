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
   StartUpPosition =   3  'Windows 기본값
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
'여러분의 서비스의 동일한 데이터 뿐만아니라 SOAP Envelope 를 빌드하여 사용할 수 있는 어떠한 일정한 것을 정의 해야 한다.
Private Const ENC = "http://schemas.xmlsoap.org/soap/encoding/"
Private Const XSI = "http://www.w3.org/1999/XMLSchema-instance"
Private Const XSD = "http://www.w3.org/1999/XMLSchema"

Private Sub Command1_Click()
''여러분의 서비스의 동일한 데이터 뿐만아니라 SOAP Envelope 를 빌드하여 사용할 수 있는 어떠한 일정한 것을 정의 해야 한다.
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
'여러분의 SOAP Connector, Serializer and Reader 초기화 해야 한다. 이 커넥터는 HTTP 커넥션으로 다룰 것이다. Serializer 는 여러분이 SOAP Envelope를 빌드 하는것을 도와 줄 것이며 Reader는 여러분이 그 결과를 사용할 수 있게 도와준다.
        Dim Connector As SoapConnector30
        Dim Serializer As SoapSerializer30
        Dim Reader As SoapReader30
        
        Set Connector = New HttpConnector30
        Set Serializer = New SoapSerializer30
        Set Reader = New SoapReader30

'사전에 SOAP Server의 커넥터를 준비 해야 한다. "SoapAction" 데이타는 서버사이드에서는 당연하지 않지만, 그 내용은 그 어떠한 모든 것을 할 수 있다. 이것은 여러분이 SOAP 메시지를 읽고 디버깅할 때 좀 더 쉽게 확인할 수 있는 URI과 메소드 이름을 세트 하는 좋은 생각이다.
        Connector.Property("EndPointURL") = url
        Call Connector.Connect
        Connector.Property("SoapAction") = URI & "#" & Method
        Call Connector.BeginMessage

'여러분의 Serialize와 Connector를 관련시키다.
        Serializer.Init Connector.InputStream

'SOAP Envelope를 시작하고 Encoding 과 XML -Schema를 명시한다.
        Serializer.StartEnvelope , ENC
        Serializer.SoapNamespace "xsi", XSI
        Serializer.SoapNamespace "SOAP-ENC", ENC
        Serializer.SoapNamespace "xsd", XSD

'메시지의 바디를 시작한다. 루트 원소는 항상 서비스 URI와 메소드여야 한다.
        Serializer.StartBody
        Serializer.StartElement Method, URI, , "method"

'루트 원소에 차일드는 각각의 메소드 파라미터에 대해 적어야 한다.
        Serializer.StartElement "Author"
        Serializer.SoapAttribute "type", , "xsd:string", "xsi"
        Serializer.WriteString "Wilde, Oscar"
        Serializer.EndElement

'루트 원소의 끝에 body와 envelope를 위치한다.
        Serializer.EndElement
        Serializer.EndBody
        Serializer.EndEnvelope

'보냈다면 메시지 끝을 종료 해야 한다.
        Connector.EndMessage

'리더안에 결과를 로드 하라.
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
'  defWSDL = "D:\프로젝트\수정구보건소\SOAP\공공보건\보건정보 Sample\WebService.asmx"
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
''wscript.echo SoapClient.AddNumbers(2, 3) '웹서비스에 정의된 메소드 호출
'
''Print low - Level; interface; 예제
'Dim Serializer As SoapSerializer30 '전송할 데이터를 SOAP XML형태로
'Dim Reader As SoapReader30 '받은 데이터를 XML 형태로
'
''Set Connector = New HttpConnector30 '해당 주소로 연결
'
''Connector.Property("EndPointURL") = "http://www.soapclient.com/interop/InteropB.wsdl" '"http://www.xxxx.com/webservice.php"
''Connector.Connect
''Connector.Property("SoapAction") = "uri:" & Method
''Connector.BeginMessage
'
'Set Serializer = New SoapSerializer30
''Serializer.Init Connector.InputStream
'
''MsgBox ("SOAP 통신 데이터생성")
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
''MsgBox ("SOAP 통신 결과 출력")
'
'Set Reader = New SoapReader30
'Reader.Load Connector.OutputStream
'Text1.Text = Reader.Body.xml
'MsgBox Reader.Body.xml

End Sub
