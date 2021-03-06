unit U_Server;

interface
  uses Windows, Classes, SysUtils, Variants, Dialogs;

const
  HCODE = 'C0501';  //보건소코드
  ICODE = 'C1';  //생화학장비

function OrderCall(BarCode:string):string;
function WorkListCall:string;
function UploadCall(BarCode:string; vExamCode, vResult:Variant; var SvrMsg:string):integer;

implementation

uses WebService, IdCoderMIME;

function OrderCall(BarCode:string):string;
var
  WSS: WebServiceSoap;
  hl7In: WideString;
  IdEncoderMIME1: TIdEncoderMIME;
  IdDecoderMIME1: TIdDecoderMIME;
begin
  Result:= '';
  IdEncoderMIME1:= TIdEncoderMIME.Create(nil);
  IdDecoderMIME1:= TIdDecoderMIME.Create(nil);
  IdDecoderMIME1.FillChar:= #128;
  try
      try
      hl7In:=IdEncoderMIME1.Encode('MSH|^~\&|HL7|MMS|||1||ORU^R01|1a082e2:10e59b48c04:-2cf9:27695009|P|2.3||||||8859/1'+#13+
                                   'PID|||'+BarCode+'^'+ICode+'^^DefaultDomain^PI'+#13+
                                   'PV1||E|'+HCODE+#13+
                                   'OBR|1||||||1');
      WSS:= GetWebServiceSoap(false, '', nil);
      Result:= IdDecoderMIME1.DecodeString( WSS.New_SelectOrder(hl7In));
      except
          on e:exception do begin
              ShowMessage(e.Message);
              exit;
          end;
      end;
  finally
      IdEncoderMIME1.Free;
      IdDecoderMIME1.Free;
  end;
end;

function WorkListCall:string;
var
  WSS: WebServiceSoap;
  hl7In, sl7Out: WideString;
  IdEncoderMIME1: TIdEncoderMIME;
  IdDecoderMIME1: TIdDecoderMIME;
begin
  Result:= '';
  IdEncoderMIME1:= TIdEncoderMIME.Create(nil);
  IdDecoderMIME1:= TIdDecoderMIME.Create(nil);
  IdDecoderMIME1.FillChar:= #128;

  try
      try
      hl7In:=IdEncoderMIME1.Encode('MSH|^~\&|HL7|MMS|||1||ORU^R01|1a082e2:10e59b48c04:-2cf9:27695009|P|2.3||||||8859/1'+#13+
                                   'PID|||^'+ICode+'^'+HCODE+'00001^DefaultDomain^PI'+#13+
                                   'PV1||E|'+HCODE+#13+
                                   'OBR|1||||||1');
      WSS:= GetWebServiceSoap(false, '', nil);
      sl7Out:= IdDecoderMIME1.DecodeString( WSS.MdbOrderList(hl7In));
      Result:= sl7Out;
      except
          on e:exception do begin
                ShowMessage(e.Message);
                exit;
          end;
      end;
  finally
      IdEncoderMIME1.Free;
      IdDecoderMIME1.Free;
  end;

end;

function UploadCall(BarCode:string; vExamCode, vResult:Variant; var SvrMsg:string):integer;
var
  WSS: WebServiceSoap;
  hl7In: WideString;
  IdEncoderMIME1: TIdEncoderMIME;
  UpDateTime:string;
  sHeader, sMid, sEnd:string;
  i:integer;
begin
  Result:= -1;
  SvrMsg:='';

  UpDateTime:= FormatDateTime('yyyymmddhhnnss', now);

  sHeader:= 'MSH|^~\&|HL7|MMS|||'+UpDateTime+'||ORU^R01|1a082e2:10e59b48c04:-2cf9:27695009|P|2.3||||||8859/1'+#13+
            'PID|||'+BarCode+'^'+ICode+'^'+HCODE+'00001^^^DefaultDomain^PI'+#13+
            'PV1||E|'+HCODE+#13+
            'OBR|1||||||'+UpDateTime+#13;

  sEnd:= 'OBR|1||||||1';
  sMid:= '';

  if VarArrayHighBound(vResult, 1) < 0 then
      exit;
  
  for I := VarArrayLowBound(vResult, 1) to VarArrayHighBound(vResult, 1) do
      sMid:= sMid + 'OBX|'+IntToStr(i+1)+'|ST|'+vExamCode[i]+'||'+vResult[i]+'||||||R'+#13;

  IdEncoderMIME1:= TIdEncoderMIME.Create(nil);
  try
      try
      hl7In:= IdEncoderMIME1.Encode(sHeader + sMid + sEnd);

      WSS:= GetWebServiceSoap(false, '', nil);
      Result:= WSS.UpdateRst(hl7In);
      except
          on e:exception do begin
              SvrMsg:= e.Message;
              exit;
          end;
      end;
  finally
      IdEncoderMIME1.Free;
  end;

end;

end.
