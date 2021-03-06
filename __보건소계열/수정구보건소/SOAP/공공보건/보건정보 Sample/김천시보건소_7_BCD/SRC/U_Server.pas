unit U_Server;

interface
  uses Windows, Classes, SysUtils, Variants, Dialogs;

const
  WSDL_DLL_NAME = 'CALL_WSDL.dll';
  HCODE = 'N0626';  //보건소코드
  ICODE = 'C';     //생화학장비

function OrderCall(BarCode:string):string;
function WorkListCall:string;
function UploadCall(BarCode:string; vExamCode, vResult:Variant; var ErrMsg:string):integer;

implementation

uses GlobalVar, U_Main;

  function DownLoadOrder_BarCode_Dll(const SvrStr:string):boolean; external WSDL_DLL_NAME;
  function DownLoadOrder_WorkList_Dll(const SvrStr:string):boolean; external WSDL_DLL_NAME;
  function UploadResult_Dll(const SvrStr:string):boolean; external WSDL_DLL_NAME;
  function GetErrorMessage:string; external WSDL_DLL_NAME;
  function GetDownLoadData:string; external WSDL_DLL_NAME;
  function GetUploadCount:integer; external WSDL_DLL_NAME;


function OrderCall(BarCode:string):string;
var
  SvrStr:string;
begin
  Result:='';
  SvrStr:= 'MSH|^~\&|HL7|MMS|||1||ORU^R01|1a082e2:10e59b48c04:-2cf9:27695009|P|2.3||||||8859/1'+#13+
           'PID|||'+BarCode+'^'+ICODE+'^^DefaultDomain^PI'+#13+
           'PV1||E|'+HCODE+#13+
           'OBR|1||||||1';
  try
      if DownLoadOrder_BarCode_Dll(SvrStr) then
          Result:= GetDownLoadData
      else
          TGlobal.SvrError:= GetErrorMessage;
  except
      TGlobal.SvrError:= GetErrorMessage;
  end;

end;

function WorkListCall:string;
var
  SvrStr:string;
begin
  Result:='';
  SvrStr:= 'MSH|^~\&|HL7|MMS|||1||ORU^R01|1a082e2:10e59b48c04:-2cf9:27695009|P|2.3||||||8859/1'+#13+
           'PID|||^'+ICODE+'^'+HCODE+'00001^DefaultDomain^PI'+#13+
           'PV1||E|'+HCODE+#13+
           'OBR|1||||||1';

  if DownLoadOrder_WorkList_Dll(SvrStr) then
      Result:= GetDownLoadData
  else
      TGlobal.SvrError:= GetErrorMessage;
end;

function UploadCall(BarCode:string; vExamCode, vResult:Variant; var ErrMsg:string):integer;
var
  sHeader, sMid, sEnd:string;
  SvrStr, UpDateTime:string;
  i:integer;
begin
  Result:= -1;

  UpDateTime:= FormatDateTime('yyyymmddhhnnss', now);

  sHeader:= 'MSH|^~\&|HL7|MMS|||'+UpDateTime+'||ORU^R01|1a082e2:10e59b48c04:-2cf9:27695009|P|2.3||||||8859/1'+#13+
            'PID|||'+BarCode+'^'+ICODE+'^'+HCODE+'00001^^^DefaultDomain^PI'+#13+
            'PV1||E|'+HCODE+#13+
            'OBR|1||||||'+UpDateTime+#13;

  sEnd:= 'OBR|1||||||1';

  if VarArrayHighBound(vResult, 1) < 0 then begin
      ErrMsg:= '결과가 0건입니다';
      exit;
  end;

  sMid:='';
  for I := VarArrayLowBound(vResult, 1) to VarArrayHighBound(vResult, 1) do
      sMid:= sMid + 'OBX|'+IntToStr(i+1)+'|ST|'+vExamCode[i]+'||'+vResult[i]+'||||||R'+#13;

  SvrStr:= sHeader + sMid + sEnd;

  if F_Main.DEBUG1.Checked then
      TGlobal.LogMsg:= SvrStr;

  if UploadResult_Dll(SvrStr) then
      Result:= GetUploadCount
  else
      ErrMsg:= GetErrorMessage;

end;

end.
