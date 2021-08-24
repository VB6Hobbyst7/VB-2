{
  에러코드
  Operator Id Length Error(Not 6): 'U'
  Operator Id No Error(1:간호사, 2:검사실, 3,4:계약직, 5,6:인턴 ,7:기타):'E'
  BarCode Length <> 11 or 8 : 'N'
  errPatient:P;
  Upload Err: 'X'
}
unit U_IFClass;

interface

uses SysUtils, Classes, variants, GlobalVar, Dialogs;

type
  TIfMaster = Class(TObject)
    FExamDate,
    FExamTime,
    FExamSeq,
    FBarCode,
    FPatId,
    FPatNm,
    FAcptDt,
    FAcptNo,
    FAge,FSex,
    FLoc,FFlag,
    FICode,
    FResult,
    FUserID,
    FIfCode,
    FAbbr,
    FExamCode,
    FQCYN,
    FPOCTYN,
    FLotNo,
    FUpState,
    SvrMsg,
    FOrdState:string;
    FRefLow:double;
    FRefHigh:double;
    constructor Create;
  end;

type
  TH7180If = Class(TIfMaster)
  private
    FBarCode:string;
    procedure SetBarCode(const Value: string);
  public
    slIfCode:TStringList;
    slExCode:TStringList;
    slIfCode_Down:TStringList;
    slResIfCode:TStringList;
    slResExCode:TStringList;
    slResult:TStringList;
    constructor Create;
    destructor Destroy; override;
    procedure DownLoadOrder;
    function MakeOrderStr:string;
    property BarCode: string read FBarCode write SetBarCode;
  end;

implementation

uses SetDataBase, U_DM, U_CodeInfo;


{ TIfMaster }

constructor TIfMaster.Create;
begin
  FUpState:= 'N';
  FOrdState:= 'N';
  FQCYN:='N';
end;

{ TH7180If }

constructor TH7180If.Create;
begin
  inherited Create;
  slIfCode:= TStringList.Create;
  slResExCode:= TStringList.Create;
  slExCode:= TStringList.Create;
  slIfCode_Down:= TStringList.Create;
  slResIfCode:= TStringList.Create;
  slResult:= TStringList.Create;
end;

destructor TH7180If.Destroy;
begin
  slIfCode.Free;
  slExCode.Free;
  slIfCode_Down.Free;
  slResIfCode.Free;
  slResult.Free;
  slResExCode.Free;
  inherited;
end;

procedure TH7180If.DownLoadOrder;
begin
  //DM.DownLoadOrder(FExamDate, FExamSeq, FBarCode, FPatId, FPatNm, FAcptNo);
  DM.DownLoadOrder(Self);
end;

function TH7180If.MakeOrderStr: string;
var
  i, nPos:integer;
  cOrdBlock:string;
begin
  Result:= '';
  cOrdBlock:='0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000';

  for i:=0 to slIfCode.Count -1 do begin
     nPos:= StrToIntDef(Trim(slIfCode.Strings[i]),0);
     if nPos > 0 then
         cOrdBlock[nPos]:='1';
  end;

  Result:= cOrdBlock;

end;

procedure TH7180If.SetBarCode(const Value: string);
begin
  if Length(Value) = 10 then
      FBarCode := '20'+Value
  else
      FBarCode := Value;
end;

end.
