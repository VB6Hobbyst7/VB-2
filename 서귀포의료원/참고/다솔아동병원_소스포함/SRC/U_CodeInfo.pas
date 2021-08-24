unit U_CodeInfo;

interface

uses ADODB, Forms, SysUtils, Graphics, Variants, Classes;

type
  TCodeInfo = Class(TObject)
    constructor Create;
    destructor destroy;   override;
  private
    procedure InitArray;
  public
    TAbbr:TStringList;
    ItemCount:integer;
    ABBRCount:integer;
    FInQuery:string;
    FaExamCode : array of string;
    FaUpCode   : array of string;
    FaIfCode   : array of string;
    FaExamName : array of string;
    FaAbbr     : array of string;
    FaDispSeq  : array of Integer;
    FaRefHigh  : array of Double;
    FaRefLow   : array of Double;
    function GetExamCode_IfCode(sIFCode:string):string;
    function GetExamCode_UpCode(sUpCode:string):string;
    function IsSetCodeOK(ECD, UPCD:string):boolean;
    function GetCodeList:integer;
    function GetAbbrCount:integer;
    function GetDispSeq(sExamCode:string): integer;
    function GetAbbr(sExamCode: string):string;
    function GetAbbr_IF(IFCD: string):string;
    function GetAbbr_Up(UPCD:string):string;
    function GetIfCode(sExamCode:string):string;
    function GetUpCode(sExamCode:string):string;
    function GetUpCode_Abbr(ABR:string):string;
    function SetCode_IfCode(sIfCode: string): boolean;
    function SetCode_UpCode(sUpCode: string): boolean;
    function SetCode_ECode(sExamCode:string):boolean;
    function CheckLowHigh_Abbr(sAbbr, sResult:string):string;
    function CheckLowHigh_Code(sCode, sResult:string):string;
    function CheckLowHigh(Index:integer; sECode, sResult:string):string;
    function GetCodeIndex(sECode:string):integer;
    function GetExamCode_Var(sIfCd:string):Variant;
  end;

var
  TCode: TCodeInfo;

implementation

uses SetDataBase, GlobalVar, DB;


constructor TCodeInfo.Create;
begin
  inherited;
  TAbbr:= TStringList.Create;
  GetCodeList;
end;

destructor TCodeInfo.destroy;
begin
  TAbbr.Free;
  FaExamCode  := nil;
  FaUpCode    := nil;
  FaExamName  := nil;
  FaAbbr      := nil;
  FaDispSeq   := nil;
  FaRefHigh    := nil;
  FaRefLow    := nil;

  inherited;
end;

function TCodeInfo.GetCodeList: integer;
var
  TSql: TQueryInfo;
  QryEx: TAdoQuery;
  i, nCount:integer;
begin
  Result:= 0;

  ABBRCount:= GetAbbrCount;

  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Application);

  try
      with TSql do begin
          AddSql(' Select * From TB_CodeInfo ');
          AddSql(' Order By DispSeq, ExamCode ');
          RCount:= LocalSelect(QryEx);
          if RCount = 0 then
              exit;

          Result:= RCount;
          ItemCount:= RCount;
          InitArray;

          with QryEx do begin
              i:=0;
              while Not Eof do begin
                  FInQuery:= FInQuery + ''''+FieldByName('ExamCode').AsString + ''',';
                  FaExamCode[i] := FieldByName('ExamCode').AsString;
                  FaUpCode[i]   := FieldByName('UpCode').AsString;
                  FaIfCode[i]   := FieldByName('IfCode').AsString;
                  FaExamName[i] := FieldByName('ExamName').ASString;
                  FaAbbr[i]     := FieldByName('Abbr').ASString;
                  FaDispSeq[i]  := FieldByName('DispSeq').ASInteger;
                  FaRefHigh[i]  := FieldByName('RefHigh').AsFloat;
                  FaRefLow[i]   := FieldByName('RefLow').AsFloat;
                  inc(i);
                  Next;
              end;
          end;
      end;

      if FInQuery <> '' then
          FInQuery:= '(' + Copy(FInQuery,1,Length(FInQuery)-1) + ')';
  finally
      QryEx.Free;
      TSql.Free;
  end;
end;

function TCodeInfo.GetUpCode(sExamCode: string): string;
var
  i:integer;
begin
  Result:= '';

  for i:=Low(FaExamCode) to High(FaExamCode) do
  begin
      if FaExamCode[i] = sExamCode then
      begin
          Result:= FaUpCode[i];
          exit;
      end;
  end
end;

procedure TCodeInfo.InitArray;
var
  i:integer;
begin
    SetLength(FaExamCode, ItemCount);
    SetLength(FaUpCode  , ItemCount);
    SetLength(FaIFCode  , ItemCount);
    SetLength(FaExamName, ItemCount);
    SetLength(FaAbbr    , ItemCount);
    SetLength(FaDispSeq , ItemCount);
    SetLength(FaRefHigh  , ItemCount);
    SetLength(FaRefLow   , ItemCount);

    for i:=0 to ItemCount -1 do
    begin
        FaExamCode [i]:='';
        FaUpCode   [i]:='';
        FaIFCode   [i]:='';
        FaExamName [i]:='';
        FaAbbr     [i]:='';
        FaDispSeq  [i]:=0;
        FaRefHigh  [i]:=0;
        FaRefLow   [i]:=0;
        
    end;
end;

function TCodeInfo.GetDispSeq(sExamCode: string): integer;
var
  i:integer;
begin
  Result:= -1;

  for i:=Low(FaExamCode) to High(FaExamCode) do
  begin
      if (FaExamCode[i] = sExamCode) then
      begin
          Result:= FaDispSeq[i];
          exit;
      end;
  end;
end;

function TCodeInfo.GetAbbr(sExamCode: string): string;
var
  i:integer;
begin
  Result:= '';

  for i:=Low(FaExamCode) to High(FaExamCode) do
  begin
      if FaExamCode[i] = sExamCode then
      begin
          Result:= FaAbbr[i];
          exit;
      end;
  end;
end;

function TCodeInfo.GetAbbrCount: integer;
var
  TSql:TQueryInfo;
  QryEx:TAdoQuery;
  nCount:integer;
begin
  Result:= 0;

  TSql:= TQueryInfo.Create;
  QryEx:= TADOQuery.Create(Application);

  TAbbr.Clear;

  try
      with TSql do begin
          Clear;
          AddSql(' Select Abbr, DispSeq From TB_CodeInfo Order By DispSeq ');
          nCount:= LocalSelect(QryEx);

          Result:= nCount;

          with QryEx do begin
              while Not Eof do begin
                  if TAbbr.IndexOf(Fields[0].AsString) < 0 then
                      TAbbr.Add(Trim(Fields[0].AsString));
                  Next;
              end;
          end;
      end;
  finally
      QryEx.Free;
      TSql.Free;
  end;

end;

function TCodeInfo.CheckLowHigh_Abbr(sAbbr, sResult: string): string;
var
  i:integer;
  sECode:string;
begin
  Result:='';
  if (sResult = '') or (sAbbr='') then
      exit;

  for i:=0 to ItemCount -1 do begin
      if sAbbr = FaAbbr[i] then begin
         sECode:= FaExamCode[i];
         Result:= CheckLowHigh(i, sECode, sResult);
         exit;
      end;
  end;

end;

function TCodeInfo.CheckLowHigh(Index:integer; sECode, sResult: string): string;
var
  dMin,dMax,dVal:double;
begin
  Result:= '';

  if UpperCase(sResult) = 'TRUE' then sResult:= '1';
  if UpperCase(sResult) = 'FALSE' then sResult:='0';
  if UpperCase(sResult) = 'NEGA' then sResult:='0';
  if UpperCase(sResult) = 'POSI' then sResult:='1';

  dVal:= StrToFloatDef(sResult,-100);
  if dVal < -99 then exit;

  dMin:= FaRefLow[Index];
  dMax:= FaRefHigh[Index];

  if dVal < dMin then
      Result:= 'L'
  else
  if dVal > dMax then
      Result:= 'H';
end;

function TCodeInfo.CheckLowHigh_Code(sCode, sResult: string): string;
var
  i:integer;
begin
  Result:='';
  if (sResult = '') or (sCode='') then
      exit;

  for i:=0 to ItemCount -1 do begin
      if sCode = FaExamCode[i] then begin
         Result:= CheckLowHigh(i, sCode, sResult);
         exit;
      end;
  end;

end;

function TCodeInfo.GetCodeIndex(sECode: string): integer;
var
  i:integer;
begin
  Result:= -1;

  for i:=Low(FaExamCode) to High(FaExamCode) do begin
      if sECode = FaExamCode[i] then begin
          Result:= i;
          exit;
      end;
  end;

end;

function TCodeInfo.SetCode_ECode(sExamCode: string): boolean;
var
  i:integer;
begin
  Result:= False;

  for i:=0 to ItemCount -1 do begin
      if sExamCode = FaExamCode[i] then begin
          Result:= True;
          exit;
      end;
  end;
end;

function TCodeInfo.SetCode_IfCode(sIfCode: string): boolean;
var
  i:integer;
begin
  Result:= False;

  for i:=0 to ItemCount -1 do begin
      if sIfCode = FaIfCode[i] then begin
          Result:= True;
          exit;
      end;
  end;
end;

function TCodeInfo.GetExamCode_Var(sIfCd: string): Variant;
var
  i, R, a:integer;
  vCode: Variant;
begin
  R:=0;

  for i:=0 to ItemCount -1 do begin
      if sIfCd = FaUpCode[i] then begin
         Inc(R)
      end;
  end;

  if R > 0 then begin
      Result:= VarArrayCreate([0, R-1], varVariant);

      a:=-1;
      for i:=0 to ItemCount -1 do begin
          if sIfCd = FaUpCode[i] then begin
             Inc(a);
             Result[a]:= FaExamCode[i];
          end;
      end;
  end
  else begin
      Result:= VarArrayCreate([0,0], varVariant);
      Result[0]:='';
  end;

end;

function TCodeInfo.IsSetCodeOK(ECD, UPCD: string): boolean;
var
  i:integer;
begin
  Result:= False;

  for i:=Low(FaExamCode) to High(FaExamCode) do
  begin
      if (FaUpCode[i] = UPCD) and (ECD = FaExamCode[i]) then
      begin
          Result:= True;
          exit;
      end;
  end

end;

function TCodeInfo.GetUpCode_Abbr(ABR: string): string;
var
  i:integer;
begin
  Result:= '';

  for i:=Low(FaAbbr) to High(FaAbbr) do
  begin
      if FaAbbr[i] = Abr then
      begin
          Result:= FaUpCode[i];
          exit;
      end;
  end

end;

function TCodeInfo.GetAbbr_IF(IFCD: string): string;
var
  i:integer;
begin
  Result:= '';

  for i:=Low(FaIFCODE) to High(FaIFCODE) do
  begin
      if FaIFCODE[i] = IFCD then
      begin
          Result:= FaAbbr[i];
          exit;
      end;
  end;

end;

function TCodeInfo.GetIfCode(sExamCode: string): string;
var
  i:integer;
begin
  Result:= '';

  for i:=Low(FaExamCode) to High(FaExamCode) do
  begin
      if FaExamCode[i] = sExamCode then
      begin
          Result:= faIfCode[i];
          exit;
      end;
  end

end;

function TCodeInfo.GetExamCode_IfCode(sIFCode: string): string;
var
  i:integer;
begin
  Result:= '';

  for i:=Low(FaExamCode) to High(FaExamCode) do
  begin
      if FaIfCode[i] = sIFCode then
      begin
          Result:= FaEXamCode[i];
          exit;
      end;
  end

end;

function TCodeInfo.GetExamCode_UpCode(sUpCode: string): string;
var
  i:integer;
begin
  Result:= '';

  for i:=Low(FaExamCode) to High(FaExamCode) do
  begin
      if FaUpCode[i] = sUpCode then
      begin
          Result:= FaEXamCode[i];
          exit;
      end;
  end

end;

function TCodeInfo.GetAbbr_Up(UPCD: string): string;
var
  i:integer;
begin
  Result:= '';

  for i:=Low(FaUPCODE) to High(FaUPCODE) do
  begin
      if FaUPCODE[i] = UPCD then
      begin
          Result:= FaAbbr[i];
          exit;
      end;
  end;
end;

function TCodeInfo.SetCode_UpCode(sUpCode: string): boolean;
var
  i:integer;
begin
  Result:= False;

  for i:=0 to ItemCount -1 do begin
      if sUpCode = FaUpCode[i] then begin
          Result:= True;
          exit;
      end;
  end;
end;

end.
