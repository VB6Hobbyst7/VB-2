// ************************************************************************ //
// The types declared in this file were generated from data read from the
// WSDL File described below:
// WSDL     : http://10.47.14.52:8009/HL7IFWebService/WebService.asmx?wsdl
//  >Import : http://10.47.14.52:8009/HL7IFWebService/WebService.asmx?wsdl:0
// Encoding : utf-8
// Version  : 1.0
// (2009-03-09 ¿ÀÈÄ 12:21:31 - * $Rev: 5154 $)
// ************************************************************************ //

unit WebService;

interface

uses InvokeRegistry, SOAPHTTPClient, Types, XSBuiltIns;

const
  IS_OPTN = $0001;
  IS_UNBD = $0002;
  IS_NLBL = $0004;
  IS_UNQL = $0008;
  IS_ATTR = $0010;
  IS_TEXT = $0020;

type

  // ************************************************************************ //
  // The following types, referred to in the WSDL document are not being represented
  // in this file. They are either aliases[@] of other types represented or were referred
  // to but never[!] declared in the document. The types from the latter category
  // typically map to predefined/known XML or Borland types; however, they could also 
  // indicate incorrect WSDL documents that failed to declare or import a schema type.
  // ************************************************************************ //
  // !:string          - "http://www.w3.org/2001/XMLSchema"
  // !:int             - "http://www.w3.org/2001/XMLSchema"
  // !:schema          - "http://www.w3.org/2001/XMLSchema"

  SelectMdbOrderListResult = class;             { "http://tempuri.org/"[Cplx] }



  // ************************************************************************ //
  // XML       : SelectMdbOrderListResult, <complexType>
  // Namespace : http://tempuri.org/
  // ************************************************************************ //
  SelectMdbOrderListResult = class(TRemotable)
  private
    Fschema: WideString;
  published
    property schema: WideString  read Fschema write Fschema;
  end;


  // ************************************************************************ //
  // Namespace : http://tempuri.org/
  // soapAction: http://tempuri.org/%operationName%
  // transport : http://schemas.xmlsoap.org/soap/http
  // style     : document
  // binding   : WebServiceSoap
  // service   : WebService
  // port      : WebServiceSoap
  // URL       : http://10.47.14.52:8009/HL7IFWebService/WebService.asmx
  // ************************************************************************ //
  WebServiceSoap = interface(IInvokable)
  ['{F790808E-B4BF-D72D-D6C0-3B8652961A63}']
    function  UpdateRst(const sHl7: WideString): Integer; stdcall;
    function  SelectRst: WideString; stdcall;
    function  SelectOrder(const sHl7: WideString): WideString; stdcall;
    function  SelectMdbOrderList(const phc_cd: WideString; const sdate: WideString; const edate: WideString): SelectMdbOrderListResult; stdcall;
    function  DeleteTestItem(const sHl7: WideString): WideString; stdcall;
    function  SelectTestItem(const sHl7: WideString): WideString; stdcall;
    function  InsertTestItem(const stringdata: WideString): WideString; stdcall;
    function  MdbOrderList(const sHl7: WideString): WideString; stdcall;
    function  New_SelectOrder(const sHl7: WideString): WideString; stdcall;
    function  User_IDSelect(const sHl7: WideString): WideString; stdcall;
  end;

function GetWebServiceSoap(UseWSDL: Boolean=System.False; Addr: string=''; HTTPRIO: THTTPRIO = nil): WebServiceSoap;


implementation
  uses SysUtils;

function GetWebServiceSoap(UseWSDL: Boolean; Addr: string; HTTPRIO: THTTPRIO): WebServiceSoap;
const
  defWSDL = 'http://10.47.14.52:8009/HL7IFWebService/WebService.asmx?wsdl';
  defURL  = 'http://10.47.14.52:8009/HL7IFWebService/WebService.asmx';
  defSvc  = 'WebService';
  defPrt  = 'WebServiceSoap';
var
  RIO: THTTPRIO;
begin
  Result := nil;
  if (Addr = '') then
  begin
    if UseWSDL then
      Addr := defWSDL
    else
      Addr := defURL;
  end;
  if HTTPRIO = nil then
    RIO := THTTPRIO.Create(nil)
  else
    RIO := HTTPRIO;
  try
    Result := (RIO as WebServiceSoap);
    if UseWSDL then
    begin
      RIO.WSDLLocation := Addr;
      RIO.Service := defSvc;
      RIO.Port := defPrt;
    end else
      RIO.URL := Addr;
  finally
    if (Result = nil) and (HTTPRIO = nil) then
      RIO.Free;
  end;
end;


initialization
  InvRegistry.RegisterInterface(TypeInfo(WebServiceSoap), 'http://tempuri.org/', 'utf-8');
  InvRegistry.RegisterDefaultSOAPAction(TypeInfo(WebServiceSoap), 'http://tempuri.org/%operationName%');
  InvRegistry.RegisterInvokeOptions(TypeInfo(WebServiceSoap), ioDocument);
  RemClassRegistry.RegisterXSClass(SelectMdbOrderListResult, 'http://tempuri.org/', 'SelectMdbOrderListResult');

end.