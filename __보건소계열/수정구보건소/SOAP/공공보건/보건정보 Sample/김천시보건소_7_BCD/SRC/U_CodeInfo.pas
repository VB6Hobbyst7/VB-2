unit U_CodeInfo;

interface

uses Classes, ADODB, Forms, SysUtils, Graphics, Variants;

type
  TCodeInfo = Class(TObject)
    constructor Create;
    destructor destroy;   override;
  private
    procedure InitArray;
  public
    ItemCount:integer;
    ABBRCount:integer;
    AbbrList:TStringList;
    Location : array of string;
    ExamCode : array of string;
    UpCode   : array of string;
    ExamName : array of string;
    Abbr     : array of string;
    DispSeq  : array of Integer;
    RefHigh  : array of Double;
    RefLow   : array of Double;
    QCYN     : array of boolean;
    function GetExamCode(sUpCode:string):string;
    function GetCodeList:integer;
    function GetAbbrCount:integer;
    function GetDispSeq(sExamCode:string): integer;
    function GetAbbr(sExamCode: string):string;
    function GetUpCode(sExamCode:string):string;
    function SetCode(sUpCode: string): boolean;
    function CheckLowHigh(sUpCode, sResult:string):string;
    function GetLowHigh(Index:integer; sResult:string):string;

  end;

  TPanelInfo = Class(TObject)
    constructor Create;
  private
    procedure InitArray;
    procedure MakeList;
  public
    ItemCount:integer;
    ICode,
    Location,
    Flag,
    PCode: array of string;
    POCT: array of boolean;
    function GetPanelCode(sLoc, sFlag:string):string;
    function GetPOCTYN(sLoc:string):boolean;
    function GetInstCode(sLoc:string):string;
  end;

  TQCInfo = Class(TObject)
    LotNo, ExamCode, UpCode:array of string;
    RefLow, RefHigh:array of double;
    ItemCount:integer;
    constructor Create;
    procedure initArray;
    procedure GetQCList;
    function SetCode(sLot,sUpCode: string): boolean;
    function CheckLowHigh(sLot, sUpCode, sResult:string):string;
    function GetLowHigh(Index:integer; sResult:string):string;
    function GetRangeData(var dLow:double; var dHigh:double; sLot,sCode:string):boolean;
    function GetExamCode(sLot,sUpCode:string):string;
  end;

var
  TCode:TCodeInfo;

implementation


uses SetDataBase, GlobalVar, DB;


constructor TCodeInfo.Create;
begin
  inherited;
  AbbrList:= TStringList.Create;
  GetCodeList;
end;

destructor TCodeInfo.destroy;
begin
  ExamCode  := nil;
  UpCode    := nil;
  ExamName  := nil;
  Abbr      := nil;
  DispSeq   := nil;
  RefHigh    := nil;
  RefLow    := nil;

  AbbrList.Free;

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
      with TSql do
      begin
          AddSql(' Select Distinct ExamCode, IFCode, ExamName, Abbr, DispSeq, RefLow, RefHigh ');
          AddSql(' From TB_Code  ');
          AddSql(' Order By DispSeq ');
          RCount:= LocalSelect(QryEx);
          if RCount = 0 then
              exit;

          Result:= RCount;
          ItemCount:= RCount;
          InitArray;

          with QryEx do
          begin
              i:=0;
              while Not Eof do
              begin
                  ExamCode[i] := FieldByName('ExamCode').AsString;
                  UpCode[i]   := FieldByName('IFCode').AsString;
                  ExamName[i] := FieldByName('ExamName').ASString;
                  Abbr[i]     := FieldByName('Abbr').ASString;
                  DispSeq[i]  := FieldByName('DispSeq').ASInteger;
                  RefHigh[i]  := FieldByName('RefHigh').AsFloat;
                  RefLow[i]   := FieldByName('RefLow').AsFloat;

                  inc(i);
                  Next;
              end;
          end;
      end;
  finally
      QryEx.Free;
      TSql.Free;
  end;
end;

function TCodeInfo.GetExamCode(sUpCode: string): string;
var
  i:integer;
begin
  Result:= '';

  for i:=Low(ExamCode) to High(ExamCode) do
  begin
      if UpCode[i] = sUpCode then
      begin
          Result:= EXamCode[i];
          exit;
      end;
  end

end;

function TCodeInfo.GetUpCode(sExamCode: string): string;
var
  i:integer;
begin
  Result:= '';

  for i:=Low(ExamCode) to High(ExamCode) do
  begin
      if ExamCode[i] = sExamCode then
      begin
          Result:= UpCode[i];
          exit;
      end;
  end
end;

procedure TCodeInfo.InitArray;
var
  i:integer;
begin
    SetLength(ExamCode, ItemCount);
    SetLength(UpCode  , ItemCount);
    SetLength(ExamName, ItemCount);
    SetLength(Abbr    , ItemCount);
    SetLength(DispSeq , ItemCount);
    SetLength(RefHigh  , ItemCount);
    SetLength(RefLow   , ItemCount);

    for i:=0 to ItemCount -1 do
    begin
        ExamCode [i]:='';
        UpCode   [i]:='';
        ExamName [i]:='';
        Abbr     [i]:='';
        DispSeq  [i]:=0;
        RefHigh  [i]:=0;
        RefLow   [i]:=0;
    end;
end;

function TCodeInfo.GetDispSeq(sExamCode: string): integer;
var
  i:integer;
begin
  Result:= -1;

  for i:=Low(ExamCode) to High(ExamCode) do
  begin
      if (ExamCode[i] = sExamCode) then
      begin
          Result:= DispSeq[i];
          exit;
      end;
  end;
end;

function TCodeInfo.GetAbbr(sExamCode: string): string;
var
  i:integer;
begin
  Result:= '';

  for i:=Low(ExamCode) to High(ExamCode) do
  begin
      if ExamCode[i] = sExamCode then
      begin
          Result:= Abbr[i];
          exit;
      end;
  end;
end;

function TCodeInfo.GetAbbrCount: integer;
var
  TSql:TQueryInfo;
  QryEx:TAdoQuery;
  i, nCount:integer;
begin
  Result:= 0;

  TSql:= TQueryInfo.Create;
  QryEx:= TADOQuery.Create(Application);

  AbbrList.Clear;

  try
      with TSql do begin
          Clear;
          AddSql(' Select Distinct Abbr, DispSeq From TB_Code Order By DispSeq ');
          nCount:= LocalSelect(QryEx);

          Result:= nCount;

          with QryEx do begin
              while Not Eof do begin
                  if AbbrList.IndexOf(Fields[0].AsString) < 0 then
                      AbbrList.Add(Trim(Fields[0].AsString));
                  Next;
              end;
          end;
      end;
  finally
      QryEx.Free;
      TSql.Free;
  end;

end;

function TCodeInfo.SetCode(sUpCode: string): boolean;
var
  i:integer;
begin
  Result:= False;

  for i:=0 to ItemCount -1 do begin
      if sUpCode = UpCode[i] then begin
          Result:= True;
          exit;
      end;
  end;
end;

function TCodeInfo.GetLowHigh(Index:integer; sResult: string): string;
var
  dMin,dMax,dVal:double;
begin
  Result:= '';

  dVal:= StrToFloatDef(sResult,-100);
  if dVal < -99 then exit;

  dMin:= RefLow[Index];
  dMax:= RefHigh[Index];

  if dVal < dMin then
      Result:= 'L'
  else
  if dVal > dMax then
      Result:= 'H';
end;

{ TPanelInfo }

constructor TPanelInfo.Create;
begin
  MakeList;
end;

function TPanelInfo.GetInstCode(sLoc: string): string;
var
  i:integer;
begin
  Result:='';

  for i:=0 to ItemCount -1 do begin
      if (sLoc = Location[i]) then begin
          Result:= ICode[i];
          exit;
      end;
  end;

end;

function TPanelInfo.GetPanelCode(sLoc, sFlag: string): string;
var
  i:integer;
begin
  Result:='';

  for i:=0 to ItemCount -1 do begin
      if (sLoc = Location[i]) and (sFlag = Flag[i]) then begin
          Result:= PCode[i];
      end;
  end;

end;

function TPanelInfo.GetPOCTYN(sLoc: string): Boolean;
var
  i:integer;
begin
  Result:=False;

  for i:=0 to ItemCount -1 do begin
      if (sLoc = Location[i]) then begin
          Result:= POCT[i];
      end;
  end;

end;

procedure TPanelInfo.InitArray;
var
  i:integer;
begin
  SetLength(ICode, ItemCount);
  SetLength(Location, ItemCount);
  SetLength(Flag, ItemCount);
  SetLength(PCode, ItemCount);
  SetLength(POCT, ItemCount);

  for i:=0 to ItemCount -1 do begin
      ICode[i]:='';
      Location[i]:='';
      Flag[i]:='';
      PCode[i]:='';
      POCT[i]:=False;
  end;

end;

procedure TPanelInfo.MakeList;
var
  TSql: TQueryInfo;
  QryEx: TAdoQuery;
  i, nCount:integer;
begin
  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Application);

  try
      with TSql do
      begin
          AddSql(' Select F.ICode, F.Flag, F.PCode, I.Location, I.POCT ');
          AddSql(' From TB_Inst I Left Join TB_Code_Flag F on (F.ICode=I.ICode) ');
          AddSql(' Order by I.DispSeq ');

          RCount:= LocalSelect(QryEx);
          if RCount = 0 then
              exit;

          ItemCount:= RCount;
          InitArray;

          with QryEx do
          begin
              i:=0;
              while Not Eof do
              begin
                  ICode[i]   := FieldByName('ICode').AsString;
                  Location[i]:= FieldByName('Location').AsString;
                  PCode[i]   := FieldByName('PCode').ASString;
                  if FieldByName('Flag').AsString = 'N' then
                      Flag[i]    := ''
                  else
                      Flag[i]:= FieldByName('Flag').AsString;

                  POCT[i]    := FieldByName('POCT').AsBoolean;
                  inc(i);
                  Next;
              end;
          end;
      end;
  finally
      QryEx.Free;
      TSql.Free;
  end;

end;

function TCodeInfo.CheckLowHigh(sUpCode, sResult: string): string;
var
  i:integer;
begin
  Result:='';
  if (sResult = '') or (sUpCode='') then
      exit;

  for i:=0 to ItemCount -1 do begin
      if sUpCode = UpCode[i]  then begin
         Result:= GetLowHigh(i, sResult);
         exit;
      end;
  end;

end;

{ TQCInfo }

function TQCInfo.GetLowHigh(Index: integer; sResult: string): string;
var
  dMin,dMax,dVal:double;
begin
  Result:= '';

  dVal:= StrToFloatDef(sResult,-100);
  if dVal < -99 then exit;

  dMin:= RefLow[Index];
  dMax:= RefHigh[Index];

  if dVal < dMin then
      Result:= 'L'
  else
  if dVal > dMax then
      Result:= 'H';
end;

function TQCInfo.CheckLowHigh(sLot, sUpCode, sResult: string): string;
var
  i:integer;
begin
  Result:='';
  if (sLot = '') or (sUpCode='') then
      exit;

  for i:=0 to ItemCount -1 do begin
      if (sLot = LotNo[i]) and (sUpCode = UpCode[i]) then begin
         Result:= GetLowHigh(i, sResult);
         exit;
      end;
  end;

end;

constructor TQCInfo.Create;
begin
  GetQCList;
end;

function TQCInfo.GetExamCode(sLot,sUpCode: string): string;
var
  i:integer;
begin
  Result:= '';

  for i:=Low(ExamCode) to High(ExamCode) do begin
      if ( LotNo[i] = sLot ) and ( UpCode[i] = sUpCode ) then begin
          Result:= EXamCode[i];
          exit;
      end;
  end;

end;

procedure TQCInfo.GetQCList;
var
  TSql: TQueryInfo;
  QryEx: TAdoQuery;
  i, nCount:integer;
begin
  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Application);

  try
      with TSql do
      begin
          AddSql(' Select LotNo, ExamCode, UpCode, RefLow, RefHigh ');
          AddSql(' From TB_QC  ');
          AddSql(' Order By LotNo, ExamCode ');

          RCount:= LocalSelect(QryEx);
          if RCount = 0 then
              exit;

          ItemCount:= RCount;
          InitArray;

          with QryEx do
          begin
              i:=0;
              while Not Eof do
              begin
                  LotNo[i]    := FieldByName('LotNo').AsString;
                  ExamCode[i] := FieldByName('ExamCode').AsString;
                  UpCode[i]   := FieldByName('UpCode').AsString;
                  RefHigh[i]  := FieldByName('RefHigh').AsFloat;
                  RefLow[i]   := FieldByName('RefLow').AsFloat;

                  inc(i);
                  Next;
              end;
          end;
      end;
  finally
      QryEx.Free;
      TSql.Free;
  end;
end;

function TQCInfo.GetRangeData(var dLow, dHigh: double; sLot,
  sCode: string): boolean;
var
  i:integer;
begin
  Result:= False;
  dLow:=0; dHigh:=0;

  for i:=0 to ItemCount -1 do begin
      if (sLot = LotNo[i]) and (sCode = ExamCode[i]) then begin
          Result:= True;
          dLow:= RefLow[i];
          dHigh:= RefHigh[i];
          exit;
      end;
  end;
end;

procedure TQCInfo.initArray;
var
  i:integer;
begin
  SetLength(LotNo, ItemCount);
  SetLength(ExamCode, ItemCount);
  SetLength(UpCode, ItemCount);
  SetLength(RefLow, ItemCount);
  SetLength(RefHigh, ItemCount);

  for i:=0 to ItemCount -1 do begin
      LotNo[i]:='';
      ExamCode[i]:='';
      UpCode[i]:='';
      RefLow[i]:=0;
      RefHigh[i]:=0;
  end;

end;

function TQCInfo.SetCode(sLot,sUpCode: string): boolean;
var
  i:integer;
begin
  Result:= False;

  for i:=0 to ItemCount -1 do begin
      if ( sLot = LotNo[i] ) and ( sUpCode = UpCode[i] ) then begin
          Result:= True;
          exit;
      end;
  end;
end;

end.
