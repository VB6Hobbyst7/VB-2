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
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1095
      Left            =   420
      TabIndex        =   1
      Top             =   1500
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   915
      Left            =   390
      TabIndex        =   0
      Top             =   480
      Width           =   2955
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''Private Sub Command1_Click()
''Dim strHL7Msg As String
''
'''Dim aa As d
'''Set aa = CreateObject("webservice_tlb")
''
''
'''strHL7Msg = aa.FetchOrder_All("B01", "20100616", "20100616", "1")
''
''
'''aa
''
'''call_wdsl
''
'''Call SoapOpen
''
'''Dim oSOAP As New SoapClient30
'''Dim Client As New SoapClient30
'''Dim send As Long
'''Dim objFile2 As Object
'''Dim d
'''Dim result
''
'''Set oSOAP = New SoapClient30
'''oSOAP.ClientProperty("ServerHTTPRequest") = True
''
'''oSOAP.mssoapinit ("http://10.47.14.52:8009/HL7IFWebService/WebService.asmx?wsdl")
''
''
'''call_wdsl
''
''                strHL7Msg = "MSH|^~\&|HL7|MMS|||1||ORU^R01|1a082e2:10e59b48c04:-2cf9:27695009|P|2.3||||||8859/1" & vbCr
''    strHL7Msg = strHL7Msg & "PID|||^" & "asas" & "^" & "1212" & "00001^DefaultDomain^PI" & vbCr
''    strHL7Msg = strHL7Msg & "PV1||E|" & "asas" & vbCr
''    strHL7Msg = strHL7Msg & "OBR|1||||||1" & vbCr
''
'''oSOAP.mssoapinit ("http://10.47.14.52:8009/HL7IFWebService/WebService.asmx?wsdl")
''
'''Call DownLoadOrder_WorkList_DLL("http://10.47.14.52:8009/HL7IFWebService/WebService.asmx?wsdl")
''Dim rv As Boolean
''
''rv = DownLoadOrder_WorkList_Dll(strHL7Msg)
''
'''dce_setenv&
'''a = New_SelectOrder.mdborderliststrHL7Msg(strHL7Msg)
''
''End Sub

Sub SoapOpen()
Dim oSOAP As New SoapClient30
Dim Client As New SoapClient30
Dim send As Long
Dim objFile2 As Object
Dim d
Dim result

Set oSOAP = New SoapClient30
'oSOAP.ClientProperty("ServerHTTPRequest") = True

oSOAP.mssoapinit ("http://10.47.14.52:8009/HL7IFWebService/WebService.asmx?wsdl")


'result = oSOAP.New_SelectOrder("1212")





'    Set oSOAP = New SoapClient30
'
'    oSOAP.ClientProperty("ServerHTTPRequest") = True
'    'oSOAP.mssoapinit ("http://10.47.14.52:8009/HL7IFWebService/WebService.asmx")
'                     '   http://10.47.14.52:8009/HL7IFWebService/WebService.asmx?wsdl
''    oSOAP.mssoapinit ("C:\Users\Administrator\Desktop\SOAP\공공보건\보건정보 Sample\WebService.wsdl")
''    oSOAP.mssoapinit ("http://10.47.14.52:8009/HL7IFWebService/WebService.asmx?wsdl")
'
'Dim strFile, strFile2
'
'    'strFile = objFile.GetFileName("E:\soapTest\Test.wsdl")
'    strFile = "C:\Users\Administrator\Desktop\SOAP\공공보건\보건정보 Sample\WebService.wsdl"
'    strFile2 = "C:\Users\Administrator\Desktop\SOAP\공공보건\보건정보 Sample\123.txt"
'
''Dim Client As New SoapClient30
'
'
'    Client.ClientProperty("ServerHTTPRequest") = True
'    Client.mssoapinit strFile
''    Client.ConnectorProperty("EndPointURL") = WSDL(url)
'    Client.ConnectorProperty("EndPointURL") = "http://10.47.14.52:8009/HL7IFWebService/WebService.asmx"
'    Client.ConnectorProperty("AuthUser") = "test"
'    Client.ConnectorProperty("AuthPassword") = "test"
'
'
'
''defWSDL = 'http://10.47.14.52:8009/HL7IFWebService/WebService.asmx?wsdl';
''defURL  = 'http://10.47.14.52:8009/HL7IFWebService/WebService.asmx';
''defSvc  = 'WebService';
''defPrt  = 'WebServiceSoap';
'
''    oSOAP.MSSoapInit2("http://10.47.14.52:8009/HL7IFWebService/WebService.asmx",,,,)
'    'result = oSOAP.AuthUser(strUserID, strPassword)
'    'result = oSOAP.SendSMS(smsID, hashValue, senderPhone, receivephone, smsContent)
'
''oSOAP.FaultString

End Sub
 
 
Private Sub Command1_Click()
    Dim SOAPClient As SoapClient30      'SOAPClient
    Set SOAPClient = New SoapClient30
    
    On Error GoTo SOAPError
    
    Dim defWSDL, defURL, defSvc, defPrt As String
    Dim aa As String
    
    defWSDL = "http://10.47.14.52:8009/HL7IFWebService/WebService.asmx?wsdl"
    defURL = "http://10.47.14.52:8009/HL7IFWebService/WebService.asmx"
    defSvc = "WebService"
    defPrt = "WebServiceSoap"

    Call SOAPClient.mssoapinit(defWSDL)
    
    aa = SOAPClient.MdbOrderList("1234567890")
    
    MsgBox "Success"
    Set SOAPClient = Nothing

Exit Sub

SOAPError:
    MsgBox SOAPClient.FaultString + vbCrLf + SOAPClient.Detail, vbOKOnly, "SOAP Error"
    MsgBox Err.Description

End Sub

'Function B64Decode(ByVal EncodeData As String) As Byte
'
'Dim Data()          As Byte
'Dim objEncode       As Base64
'
'    Set objEncode = New Base64
'
'    If IsNull(EncodeData) Or Len(Trim$(EncodeData)) < 10 Then
'        B64Decode = ""
'    Else
''        Data = objEncode.DecodeArr(EncodeData)
'        B64Decode = objEncode.DecodeArr(EncodeData)
'
'    End If
'
'End Function
'
'Function B64Encode(ByVal DecodeData As String) As String
''    '-- Encode
'    Set objEncode = New Base64
'
'    If IsNull(EncodeData) Or Len(Trim$(EncodeData)) < 10 Then
'        B64Encode = ""
'    Else
'        B64Encode = objEncode.EncodeArr(DecodeData)
'    End If
'
'End Function
'
'Sub mssoapinit(WSDLFile As String, ServiceName As String, Port As String, WSMLFile As String)
'
'End Sub
'
'
'Set SOAPClient = CreateObject("MSSOAP.SOAPClient")
'Call SOAPClient.mssoapinit("DocSample1.wsdl", "", "", "")
'wscript.echo SOAPClient.AddNumbers(2, 3)
'wscript.echo SOAPClient.SubtractNumbers(3, 2)

Private Sub Command2_Click()
    Dim SOAPClient As SoapClient30      'SOAPClient
    Set SOAPClient = New SoapClient30
    
    On Error GoTo SOAPError
    
    Dim defWSDL, defURL, defSvc, defPrt As String
    
    defWSDL = "http://tempuri.org/WebService.asmx?wsdl"
    defURL = "http://10.47.14.52:8009/HL7IFWebService/WebService.asmx"
    defWSDL = "http://microsoft.com/webservices/WebService.asmx?wsdl"
    defURL = "D:\프로젝트\수정구보건소\SOAP\공공보건\보건정보 Sample\.asmx"
    'D:\프로젝트\수정구보건소\SOAP\공공보건\보건정보 Sample
    defSvc = "WebService"
    defPrt = "WebServiceSoap"

    Call SOAPClient.mssoapinit(defSvc, defPrt, defWSDL)
    
    
'<WebService(Namespace:="http://microsoft.com/webservices/")> Public Class MyWebService
    ' 구현
'End Class


    MsgBox "Success"
    Set SOAPClient = Nothing

Exit Sub

SOAPError:
    MsgBox SOAPClient.FaultString + vbCrLf + SOAPClient.Detail, vbOKOnly, "SOAP Error"
    MsgBox Err.Description

End Sub

Private Sub Form_Load()

'Set sc = WScript.CreateObject("MSSOAP.SOAPClient30")
'sc.mssoapinit "http://localhost/DocSample7/DocSample7.wsdl"
'sc.HeaderHandler = WScript.CreateObject("SessionInfoClient.clientHeaderHandler")
'sc.SomeMethod "param1", "param2"


'''  Dim SOAPClient As SoapClient30
''''  Set SOAPClient = New SoapClient30
'''  On Error GoTo SOAPError
'''  SOAPClient.mssoapinit ("C:\Users\Administrator\Desktop\SOAP\공공보건\보건정보 Sample\WebService.wsdl")
'''  MsgBox Str(SOAPClient.getRate("England", "Japan")), vbOKOnly, "Exchange Rate"
'''Exit Sub
'''SOAPError:
'''  MsgBox SOAPClient.FaultString + vbCrLf + SOAPClient.Detail, vbOKOnly, "SOAP Error"
'''
'''Call SOAPClient.mssoapinit("DocSample1.wsdl", "", "", "")
'''
'''
''''Set SOAPClient = CreateObject("MSSOAP.SoapClient30")
''''Call SOAPClient.MSSoapInit("DocSample1.wsdl", "", "", "")
''''Debug.Print SOAPClient.AddNumbers(2, 3)
''''Debug.Print SOAPClient.SubtractNumbers(3, 2)
End Sub
