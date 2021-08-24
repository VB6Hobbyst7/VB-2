unit U_MasterInfo;

interface

uses SysUtils, Classes;

{
type
  TOneResultInfo = Class(TObject)
    UpCode,
    RltVal,
    RltTxt,
    ExamCode:string;
  end;

  TResultInfo = Class(TList)
  constructor create;
  destructor Destroy; override;
  private
    FPatId,
    FSex,
    FPatName,
    FExamDate,
    FSpcid:string;
    FExamSeq:integer;
    FOrderYN:boolean;
    FState:integer;
    function GetExamDate: string;
    function GetOrderYN: boolean;
    function GetSeq: integer;
    function GetSpcid: string;
    procedure SetExamDate(const Value: string);
    procedure SetOrderYN(const Value: boolean);
    procedure SetSeq(const Value: integer);
    procedure SetSpcid(const Value: string);
    function GetState: integer;
    procedure SetState(const Value: integer);
    function GetPatId: string;
    function GetSex: string;
    procedure SetPatId(const Value: string);
    procedure SetSex(const Value: string);
    function GetPatName: string;
    procedure SetPatName(const Value: string);
  public
    FExamCode:array of string;
    property ExamDate:string read GetExamDate write SetExamDate;
    property Spcid:string read GetSpcid write SetSpcid;
    property ExamSeq:integer read GetSeq write SetSeq;
    property OrderYN:boolean read GetOrderYN write SetOrderYN;
    property State:integer read GetState write SetState;
    property PatId:string read GetPatId write SetPatId;
    property Sex:string read GetSex write SetSex;
    property PatName:string read GetPatName write SetPatName;
    function MakeUploadStr:string;
    function EquilDownCodeYN(cExamCode:string):boolean;
  end;

  TMasterInfo = class(TObject)
    FSpcid :string;
    FExamDate:string;
    FExamSeq:integer;
    //FPatId:string;
    //FPatName:string;
    //FPatSex:string;
    FExamCode: array of string;
    FUpCode  : array of string;
    FIfCode  : array of string;
    FRsltVal : array of string;
    FRsltTxt : array of string;
    FOrderCount:integer;
    FUploadCount:integer;
    FOrderYN:boolean;
    FUpState:integer;
  private
    ItemCount:integer;
    ItemIndex:integer;
  public
    procedure AddItem(cIfCode, cUpCode,cExamCode, cRsltVal:string);
    procedure InitArray(cSpcid:string; nExamSeq, nCount:integer);
    function GetLastResultStr:string;
  end;
}

  TAblInfo = class(TObject)
    constructor create(cSpcid:string; nSeq:integer; dDate:TDateTime);
  private
    FSpcid:string;
    FExamDate:string;
    FExamSeq:integer;
    FItemIndex:integer;   //코드 받을때 사용.
    FItemCount:integer;
    FOrderYN:boolean;
    FState:integer;
    FDownCode: array of string;
    FUpCode  : array of string;
    FIfCode  : array of string;
    FRsltTxt : array of string;
    FQcYN:boolean;
    procedure AddItem(nindex:integer; cRsltTxt:string);
    procedure AddDownCode(nIndex:integer; cExamCode, cUpCode:string);
    function GetSpcid: string;
    procedure SetSpcid(const Value: string);
    function GetOrderYN:boolean;
    procedure SetOrderYN(const Value: boolean);
    function GetState:integer;
    procedure SetState(const Value: integer);
    procedure SetExamDate(const Value: string);
    procedure SetExamSeq(const Value: integer);
    function GetIfCode(Index: integer): string;
    procedure SetIfCode(Index: integer; const Value: string);
    function GetUpCode(Index:integer):string;
    procedure SetUpCode(Index:integer; cUpCode:string);
    function GetDown(Index:integer):string;
    procedure SetDown(Index:integer; cExamCode:string);
  public
    ViewExamTime,
    QCNo,
    PatId,
    PatName,
    InstName,
    Location,
    Sex:string;
    procedure AddResult(cUpCode, cRsltTxt:string);
    procedure AddExamCode(cExamCode, cUpCode:string);
    property Spcid:string read GetSpcid write SetSpcid;
    property OrderYN: boolean read GetOrderYN write SetOrderYN;
    property State: integer read GetState write SetState;
    property ExamDate:string read FExamDate write SetExamDate;
    property ExamSeq:integer read FExamSeq write SetExamSeq;
    function MakeResultStr:string;
    function GetDownCode(cUpCode:string):string;
    procedure SetCodeLength(nCount:integer);
    property ItemCount:integer read FItemCount;
    property ItemIndex:integer read FItemIndex;
    property UpCode[Index:integer]:string read GetUpCode write SetUpCode;
    property DownCode[Index:integer]:string read GetDown write SetDown;
    property IfCode[Index:integer]:string read GetIfCode write SetIfCode;
    property QcYN:boolean read FQcYN write FQcYN;
  end;

implementation

uses Variants, GlobalVar, U_DM;

{ TMasterInfo }

procedure TMasterInfo.AddItem(cIfCode, cUpCode, cExamCode, cRsltVal:string);
begin
  inc(ItemIndex);
  FUpCode[ItemIndex]:= cUpCode;
  FIfCode[ItemIndex]:= cIfCode;
  FRsltVal[ItemIndex]:= cRsltVal;
  FExamCode[ItemIndex]:= cExamCode;
  //FRsltTxt[ItemIndex]:= TCode.MakeResultTxt(cUpCode,cRsltVal);
end;

function TMasterInfo.GetLastResultStr: string;
begin
    if ItemCount > 0 then
        Result:=Format('%12s %10s %10s',[FSpcid, FUpCode[ItemIndex], FRsltTxt[ItemIndex]])
end;

procedure TMasterInfo.InitArray(cSpcid:string; nExamSeq, nCount:integer);
var
  i:integer;
begin
    FSpcid:= Copy(Trim(cSpcid),1,7);
    FExamDate:= FormatDateTime('yyyymmdd', now);
    FExamSeq := nExamSeq;
    FOrderYN := False;

    ItemCount:=nCount;
    ItemIndex:=-1;

    FOrderCount:=0;
    FUploadCount:=0;

    SetLength(FExamCode, nCount);
    SetLength(FUpCode, nCount);
    SetLength(FIfCode, nCount);
    SetLength(FRsltVal, nCount);
    SetLength(FRsltTxt, nCount);

    for i:=0 to ItemCount -1 do
    begin
        FExamCode[i]:='';
        FUpCode[i]  :='';
        FIfCode[i]  :='';
        FRsltVal[i] :='';
        FRsltTxt[i] :='';
    end;
end;

{ TResultInfo }

constructor TResultInfo.create;
begin
  inherited;
  OrderYN:= False;
  State:= stOrderN;
end;

destructor TResultInfo.Destroy;
var
  i:integer;
begin
  for i:=0 to Count -1 do
  begin
      TOneResultInfo(Items[i]).Free;
  end;

  inherited;
end;

function TResultInfo.EquilDownCodeYN(cExamCode: string): boolean;
var
  i:integer;
begin
  Result:= False;
  if Trim(cExamCode)='' then
      exit;
      
  for i:=0 to High(FExamCode) do
  begin
      if FExamCode[i] = cExamCode then
      begin
          Result:= True;
          exit;
      end;
  end;
end;

function TResultInfo.GetExamDate: string;
begin
  Result:= FExamDate;
end;

function TResultInfo.GetOrderYN: boolean;
begin
  Result:= FOrderYN;

end;

function TResultInfo.GetPatId: string;
begin
  Result:= FPatId;
end;

function TResultInfo.GetPatName: string;
begin
  Result:= FPatName;
end;

function TResultInfo.GetSeq: integer;
begin
  Result:= FExamSeq;
end;

function TResultInfo.GetSex: string;
begin
  Result:= FSex;
end;

function TResultInfo.GetSpcid: string;
begin
  Result:= FSpcid;
end;

function TResultInfo.GetState: integer;
begin
  Result:= FState;
end;

function TResultInfo.MakeUploadStr: string;
var
  i:integer;
  cCode, cResult, cData:string;
begin
  Result:= '';
  if Trim(Spcid) = '' then
      exit;

  for i:= 0 to Self.Count -1 do
  begin
      cCode  := TOneResultInfo(Items[i]).ExamCode;
      cResult:= TOneResultInfo(Items[i]).RltTxt;
      if Trim(cCode) <> '' then
          cData:= cData + cCode +'\'+cResult+'\'+',';
  end;

  Result:= Copy(cData, 1, Length(cData)-1);

end;

procedure TResultInfo.SetExamDate(const Value: string);
begin
  FExamDate:= Value;
end;

procedure TResultInfo.SetOrderYN(const Value: boolean);
begin
  FOrderYN:= Value;
end;

procedure TResultInfo.SetPatId(const Value: string);
begin
  FPatId:= Value;
end;

procedure TResultInfo.SetPatName(const Value: string);
begin
  FPatName:= Value;
end;

procedure TResultInfo.SetSeq(const Value: integer);
begin
  FExamSeq:= Value;
end;

procedure TResultInfo.SetSex(const Value: string);
begin
  FSex:= Value;
end;

procedure TResultInfo.SetSpcid(const Value: string);
begin
  FSpcid:= Value;
end;

procedure TResultInfo.SetState(const Value: integer);
begin
  FState:= Value;
end;

{ TAblInfo }

procedure TAblInfo.AddDownCode(nIndex: integer; cExamCode, cUpCode:string);
begin
  FDownCode[nIndex]:= cExamCode;
  FUpCode[nIndex]  := cUpCode;
  //FIfCode[nIndex]  := TCode.GetIfCode(cExamCode);
end;

procedure TAblInfo.AddExamCode(cExamCode, cUpCode: string);
begin
  if cExamCode = '' then
      exit;

  if ItemIndex < ItemCount then
      AddDownCode(ItemIndex, cExamCode, cUpCode);

  inc(FItemIndex);

end;

procedure TAblInfo.AddItem(nindex: integer; cRsltTxt:string);
begin
  FRsltTxt[nIndex]:= cRsltTxt;
end;

procedure TAblInfo.AddResult(cUpCode, cRsltTxt: string);
var
  i:integer;
begin
  if (cUpCode='') or (cRsltTxt='') then
      exit;

  for i:=0 to ItemCount - 1 do
  begin
      if FUpCode[i] = cUpCode then begin
         AddItem(i, cRsltTxt);
      end;
  end;

end;

constructor TAblInfo.create(cSpcid:string; nSeq:integer;  dDate:TDateTime);
begin
  inherited Create;

  FSpcid:= cSpcid;
  FExamDate:= FormatDateTime('yyyymmdd', dDate);
  FExamSeq:= nSeq;
  FItemIndex:=-1;
  FItemCount:=-1;
end;

function TAblInfo.GetDownCode(cUpCode: string): string;
var
  i:integer;
begin
  Result:= '';
  if cUpCode = '' then
      exit;

  //일단 다운받은 검사 코드 비교.
  for i:= Low(FUpCode) to High(FUpCode) do
  begin
      if FUpCode[i] = cUpCode then begin
          Result:= FDownCode[i];
          exit;
      end;
  end;

end;

function TAblInfo.GetDown(Index: integer): string;
begin
  Result:='';
  if Index < ItemCount then
      Result:= FDownCode[Index];
end;

function TAblInfo.GetOrderYN: boolean;
begin
  Result:= FOrderYN;
end;

function TAblInfo.GetSpcid: string;
begin
  Result:= FSpcid;
end;

function TAblInfo.GetState: integer;
begin
  Result:= FState;
end;

function TAblInfo.GetUpCode(Index: integer): string;
begin
  Result:='';
  if Index < ItemCount then
      Result:= FUpCode[Index];
end;

function TAblInfo.MakeResultStr: string;
var
  II:integer;
  cCode, cResult, cData, cUpCode:string;
begin
  Result:= '';
  if Trim(Spcid) = '' then
      exit;

  cData:='';
  for II:= Low(FDownCode) to High(FDownCode) do
  begin
      cCode  := FDownCode[II];
      cResult:= FRsltTxt[II];
      cUpCode:= FUpCode[II];

      if FQcYn then begin
          if (cUpCode = 'pH') or (cUpCode = 'pCO2') or (cUpCode = 'pO2') then
              continue;
      end
      else begin
          if (cUpCode = 'pH(T)') or (cUpCode = 'pCO2(T)') or (cUpCode = 'pO2(T)') then
              continue;
      end;

      if (cCode <> '') and (cResult<>'') then
          cData:= cData + cCode + #9 + cResult + #10;
  end;

  Result:= cData;
end;

procedure TAblInfo.SetCodeLength(nCount: integer);
begin
  SetLength(FDownCode, nCount);
  SetLength(FUpCode  , nCount);
  SetLength(FRsltTxt , nCount);
  SetLength(FIfCode  , nCount);

  FItemIndex:=0;
  FItemCount:=nCount;
end;

procedure TAblInfo.SetDown(Index: integer; cExamCode: string);
begin
  if Index < ItemCount then
      FDownCode[Index]:= cExamCode;
end;

procedure TAblInfo.SetExamDate(const Value: string);
begin
  FExamDate := Value;
end;

procedure TAblInfo.SetExamSeq(const Value: integer);
begin
  FExamSeq := Value;
end;

procedure TAblInfo.SetOrderYN(const Value: boolean);
begin
  FOrderYN := Value;
end;

procedure TAblInfo.SetSpcid(const Value: string);
begin
  FSpcid:= Value;
end;

procedure TAblInfo.SetState(const Value: integer);
begin
  FState := Value;
end;

procedure TAblInfo.SetUpCode(Index: integer; cUpCode: string);
begin
  if Index < Self.ItemCount then
      FUpCode[Index]:= cUpCode;
end;

function TAblInfo.GetIfCode(Index: integer): string;
begin
  Result:='';
  if Index < ItemCount then
      Result:= FIfCode[Index];
end;

procedure TAblInfo.SetIfCode(Index: integer; const Value: string);
begin
  if Index < Self.ItemCount then
      FIfCode[Index]:= Value;

end;

end.
