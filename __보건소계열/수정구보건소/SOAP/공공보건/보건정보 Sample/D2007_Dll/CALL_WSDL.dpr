library CALL_WSDL;

{ Important note about DLL memory management: ShareMem must be the
  first unit in your library's USES clause AND your project's (select
  Project-View Source) USES clause if your DLL exports any procedures or
  functions that pass strings as parameters or function results. This
  applies to all strings passed to and from your DLL--even those that
  are nested in records and classes. ShareMem is the interface unit to
  the BORLNDMM.DLL shared memory manager, which must be deployed along
  with your DLL. To avoid using BORLNDMM.DLL, pass string information
  using PChar or ShortString parameters. }

uses
  SysUtils,
  Classes,
  IdCoderMIME,
  IdCoder,
  IdCoder3to4,
  IdBaseComponent,
  //Dialogs,
  Variants,
  ActiveX,
  WebService in 'WebService.pas';



{$R *.res}

var
  DownData:string;
  UploadCount:integer;
  ErrorMessage:string;

function GetErrorMessage:string; export;
begin
  Result:= ErrorMessage;
end;

function GetDownLoadData:string; export;
begin
  Result:= DownData;
end;

function GetUploadCount:integer; export;
begin
  Result:= UploadCount;
end;

function DownLoadOrder_BarCode_Dll(const SvrStr:string):boolean; export;
var
  WSS: WebServiceSoap;
  hl7In, hl7Out: WideString;
  IdEncoderMIME1: TIdEncoderMIME;
  IdDecoderMIME1: TIdDecoderMIME;
begin
  Result:= False;
  CoInitialize(nil);
  ErrorMessage:='';

  try
      try
          DownData:= '';

          IdEncoderMIME1:= TIdEncoderMIME.Create(nil);
          IdDecoderMIME1:= TIdDecoderMIME.Create(nil);
          IdDecoderMIME1.FillChar:= #128;
          try
              hl7In:=IdEncoderMIME1.Encode(SvrStr);
              WSS:= GetWebServiceSoap(false, '', nil);
              hl7Out:= IdDecoderMIME1.DecodeString( WSS.New_SelectOrder(hl7In));
              DownData:= Copy(hl7Out,1,Length(hl7Out));

          finally
              IdEncoderMIME1.Free;
              IdDecoderMIME1.Free;
          end;
      except
          on e:exception do begin
              ErrorMessage:= e.Message;
              exit;
          end;
      end;

      Result:= True;

  finally
      CoUninitialize;
  end;

end;

function DownLoadOrder_WorkList_Dll(const SvrStr:string):boolean; export;
var
  WSS: WebServiceSoap;
  hl7In, hl7Out: WideString;
  IdEncoderMIME1: TIdEncoderMIME;
  IdDecoderMIME1: TIdDecoderMIME;
begin
  DownData:= '';
  Result:= False;
  ErrorMessage:='';

  CoInitialize(nil);
  try
      try
          IdEncoderMIME1:= TIdEncoderMIME.Create(nil);
          IdDecoderMIME1:= TIdDecoderMIME.Create(nil);
          IdDecoderMIME1.FillChar:= #128;
          try
              hl7In:=IdEncoderMIME1.Encode(SvrStr);
              WSS:= GetWebServiceSoap(false, '', nil);
              hl7Out:= IdDecoderMIME1.DecodeString( WSS.MdbOrderList(hl7In) );
              DownData:= Copy(hl7Out,1,Length(hl7Out));
          finally
              IdEncoderMIME1.Free;
              IdDecoderMIME1.Free;
          end;
      except
          on e:exception do begin
              ErrorMessage:= e.Message;
              exit;
          end;
      end;

      Result:= True;

  finally
      CoUninitialize;
  end;

end;

function UploadResult_Dll(const SvrStr:string):boolean; export;
var
  WSS: WebServiceSoap;
  hl7In: WideString;
  IdEncoderMIME1: TIdEncoderMIME;
begin
  UploadCount:= -1;
  ErrorMessage:='';
  Result:= False;

  CoInitialize(nil);
  try
      try
          IdEncoderMIME1:= TIdEncoderMIME.Create(nil);
          try

              hl7In:= IdEncoderMIME1.Encode(SvrStr);
              WSS:= GetWebServiceSoap(false, '', nil);
              //ShowMessage('WSS????');
              UploadCount:= WSS.UpdateRst(hl7In);
              //ShowMessage('UpLoad????');
          finally
              IdEncoderMIME1.Free;
          end;
      except
          on e:exception do begin
              //ShowMessage('?????????? ????'+e.Message);
              ErrorMessage:= e.Message;
              exit;
          end;
      end;

      if UploadCount > 0 then
          Result:= True;
  finally
    CoUninitialize;
  end;
end;

exports
    GetErrorMessage,
    GetDownLoadData,
    GetUploadCount,
    DownLoadOrder_BarCode_Dll,
    DownLoadOrder_WorkList_Dll,
    UploadResult_Dll;

begin

end.
