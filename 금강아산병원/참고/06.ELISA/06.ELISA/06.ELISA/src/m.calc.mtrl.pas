unit m.calc.mtrl;

interface

uses
  m.rawdata, m.curvefit,

  System.SysUtils, System.Classes, Spring.Collections, Spring
  ;

type
  TValueKind = (
    vkOver = 0,
    vkValid,
    vkInvalid
  );
  TValueKindHelper = record helper for TValueKind
    class function Create(const ASource: string; var AConvert: Double): TValueKind; static;
  end;

  TQualitativeResult = (
    qrNegative = 0,
    qrPositive,
    qrIndeterminate
  );
  TQualitativeResultHelper = record helper for TQualitativeResult
    function AsString: string;
  end;

  TQualitativeValue = record
  private const
    NIdxIUML = 0;
    NIdxLog10 = 1;
  private
    FSource: String;
    FValue: Double;
    FIndex: Integer;
    FKind: TValueKind;
    FIuml: Double;
    FLog10: Double;
    function GetAsDoubleString(const Index: Integer): String;
  public
    constructor Create(ASource: String; const ASrcIdx: Integer; const AFormula: TFormulaCalc);

    property Source: String read FSource;
    property Value: Double read FValue;
    property Index: Integer read FIndex;
    property Kind: TValueKind read FKind;
    property IUML: Double read FIuml;
    property Log10: Double read FLog10;
    property AsIUMLString: String index NIdxIUML read GetAsDoubleString;
    property AsLog10String: String index NIdxIUML read GetAsDoubleString;
  end;

  TSamplesCalculator = class
  private const
    NSrcValue = 0;
    NSrcIUML = 1;
    NSrcLog10 = 2;
  private
    FSources: array[mNil..mMitogen] of IList<TQualitativeValue>;
    FResults: IList<TQualitativeResult>;
  private
    function GetIumlDeltaMtz(AIdx: Integer): Double;
    function GetIumlDeltaMtzTexts(AIdx: Integer): String;
    function GetIumlDeltaTBAgPerNils(AIdx: Integer): Double;
    function GetIumlDeltaTBAg(AIdx: Integer): Double;
    function GetIumlDeltaTBAgTexts(AIdx: Integer): String;
    function GetSrcs(Mtrl: TMaterial; Idx: Integer): TQualitativeValue;
    function GetResults(AIdx: Integer): TQualitativeResult;
    function GetResultTexts(AIdx: Integer): String;
    function GetCount: Integer;
    function GetSrcIdxs(Idx: Integer): Integer;
    function GetSrcField(Mtrl: TMaterial; Idx: Integer; const Index: Integer): Double;
  protected
    function DoCalculate: Boolean; virtual; abstract;

    property IumlDeltaTBAgPerNils[AIdx: Integer]: Double read GetIumlDeltaTBAgPerNils;
  public
    constructor Create; virtual;

    procedure Clear;
    function Execute(const ASamples: TStringDynArray; const ASampleIdx: Integer; const ACurveFormula: TFormulaCalc): Boolean;
    function Exists(const AKey: TMaterial; const AIdx: Integer): Boolean;

    property Count: Integer read GetCount;
    property Srcs[Mtrl: TMaterial; Idx: Integer]: TQualitativeValue read GetSrcs; default;
    property SrcValues[Mtrl: TMaterial; Idx: Integer]: Double index NSrcValue read GetSrcField;
    property SrcIumls[Mtrl: TMaterial; Idx: Integer]: Double index NSrcIuml read GetSrcField;
    property SrcLog10s[Mtrl: TMaterial; Idx: Integer]: Double index NSrcLog10 read GetSrcField;
    property SrcIdxs[Idx: Integer]: Integer read GetSrcIdxs;
    property IumlDeltaTBAg[AIdx: Integer]: Double read GetIumlDeltaTBAg;
    property IumlDeltaTBAgTexts[AIdx: Integer]: String read GetIumlDeltaTBAgTexts;

    property IumlDeltaMtz[AIdx: Integer]: Double read GetIumlDeltaMtz;
    property IumlDeltaMtzTexts[AIdx: Integer]: String read GetIumlDeltaMtzTexts;

    property Results[AIdx: Integer]: TQualitativeResult read GetResults;
    property ResultTexts[AIdx: Integer]: String read GetResultTexts;
  end;

  TM2Calculator = class(TSamplesCalculator)
  protected
    function DoCalculate: Boolean; override;
  end;

  TM3Calculator = class(TSamplesCalculator)
  protected
    function DoCalculate: Boolean; override;
  end;

  TMtrlCalculator = class
  private
    FM3Calc: TM3Calculator;
    FM2Calc: TM2Calculator;
    FCount: Integer;
  public
    constructor Create;
    destructor Destroy; override;

    procedure Clear;
    function Execute(const ASamples: TArray<TStringDynArray>; const ACurveFormula: TFormulaCalc): Boolean;

    property Count: Integer read FCount;
    property M3: TM3Calculator read FM3Calc;
    property M2: TM2Calculator read FM2Calc;
  end;

function DoubleToString(const AValue: Double): String;
function IfThen(const ACondition: Boolean; const ATrue, AFalse: TQualitativeResult):  TQualitativeResult; overload;
function IfThen(const ACondition: Boolean; const ATrue, AFalse: TSamplesCalculator):  TSamplesCalculator; overload;

implementation

uses
  System.Math, System.Types, CodeSiteLogging, mCodeSiteHelper, Spring.SystemUtils, System.StrUtils
  ;

function DoubleToString(const AValue: Double): String;
begin
//  case CompareValue(AValue, 10.0) of
//    LessThanValue   : Result := AValue.ToString;
//
//    EqualsValue     ,
//    GreaterThanValue: Result := '> 10';
//  end;
  Result := IfThen(AValue < 10, (Trunc(AValue *100) /100).ToString(ffFixed, 5, 2, FormatSettings), '> 10');
//  Result := IfThen(AValue < 10, (SimpleRoundTo(AValue *100) /100).ToString(ffFixed, 5, 2, FormatSettings), '> 10');
//  Result := IfThen(AValue < 10, SimpleRoundTo(AValue).ToString(ffGeneral, 5, 2, FormatSettings), '> 10');
end;

function IfThen(const ACondition: Boolean; const ATrue, AFalse: TQualitativeResult):  TQualitativeResult;
begin
  if ACondition then
    Result := ATrue
  else
    Result := AFalse;
end;

function IfThen(const ACondition: Boolean; const ATrue, AFalse: TSamplesCalculator):  TSamplesCalculator;
begin
  if ACondition then
    Result := ATrue
  else
    Result := AFalse;
end;

{ TValueKindHelper }

class function TValueKindHelper.Create(const ASource: string; var AConvert: Double): TValueKind;
begin
  if ASource.Trim.ToLower.Equals('over') then
  begin
    Result := vkOver;
    AConvert := 10.;
  end
  else if TryStrToFloat(ASource, AConvert) then
    Result := vkValid
  else
    Result := vkInvalid;
end;

{ TQualitativeResultHelper }

function TQualitativeResultHelper.AsString: string;
begin
  Result := TEnum.GetName<TQualitativeResult>(Self).Substring(2).ToUpper;
end;

{ TQualitativeValue }

constructor TQualitativeValue.Create(ASource: String; const ASrcIdx: Integer; const AFormula: TFormulaCalc);
begin
  Assert(AFormula.HasValue, 'Parameter error, AFormula has not value!!');
  Assert(AFormula.Terms = ftLinearEquation, 'Parameter error, AFormula equation is not linear!!');

  FSource := ASource;
  FIndex := ASrcIdx;
  FKind := TValueKind.Create(FSource, FValue);
  case FKind of
    vkOver:
    begin
      FLog10 := 10.;
      FIuml := 10.;
    end;

    vkValid:
    begin
      FLog10 := System.Math.Log10(FValue);
      FIuml := Power(10, AFormula.bX1 * FLog10 + AFormula.cX0);
    end;

    vkInvalid: ;
  end;
end;

function TQualitativeValue.GetAsDoubleString(const Index: Integer): String;
begin
  Result := '';
  case Index of
    NIdxIUML: Result := DoubleToString(FIuml);
    NIdxLog10: Result := DoubleToString(FLog10);
  else
    Assert(False, 'Handled code not exists');
  end;
end;

{ TSamplesCalculator }

function TSamplesCalculator.Execute(const ASamples: TStringDynArray; const ASampleIdx: Integer; const ACurveFormula: TFormulaCalc): Boolean;
const
  SErMsg = 'Invalid parameter, TSamplesCalc.Add can handle a cmNil_Antigen or a cmNil_Antigen_Mitogen only.';
var
  i: Integer;
begin
  for i := 0 to Length(ASamples) -1 do
    FSources[TMaterial.Create(i)].Add(TQualitativeValue.Create(ASamples[i], ASampleIdx, ACurveFormula));
  Result := DoCalculate;
end;

procedure TSamplesCalculator.Clear;
var
  m: TMaterial;
begin
  FResults.Clear;
  for m in TMaterialAll do
    FSources[m].Clear;
end;

constructor TSamplesCalculator.Create;
var
  m: TMaterial;
begin
  FResults := TCollections.CreateList<TQualitativeResult>;
  for m in TMaterialAll do
    FSources[m] := TCollections.CreateList<TQualitativeValue>;
end;

function TSamplesCalculator.Exists(const AKey: TMaterial; const AIdx: Integer): Boolean;
begin
  Result := AIdx < FSources[AKey].Count;
end;

function TSamplesCalculator.GetIumlDeltaMtz(AIdx: Integer): Double;
begin
  Result := 0.0;
  case FSources[mMitogen][AIdx].Kind of
    vkOver: Result := 10.;
    vkValid: Result := FSources[mMitogen][AIdx].IUML - FSources[mNil][AIdx].IUML;
  end;
end;

function TSamplesCalculator.GetIumlDeltaMtzTexts(AIdx: Integer): String;
begin
  Result := DoubleToString(IumlDeltaMtz[AIdx]);
end;

function TSamplesCalculator.GetIumlDeltaTBAgPerNils(AIdx: Integer): Double;
begin
  Result := IumlDeltaTBAg[AIdx] / SrcValues[mNil, AIdx] * 100;
end;

function TSamplesCalculator.GetIumlDeltaTBAg(AIdx: Integer): Double;
begin
  Result := 0.0;
  case FSources[mTBAg][AIdx].Kind of
    vkOver: Result := 10.;
    vkValid: Result := FSources[mTBAg][AIdx].IUML - FSources[mNil][AIdx].IUML;
  end;
end;

function TSamplesCalculator.GetIumlDeltaTBAgTexts(AIdx: Integer): String;
begin
  Result := DoubleToString(IumlDeltaTBAg[AIdx]);
end;

function TSamplesCalculator.GetCount: Integer;
begin
  Result := FSources[mNil].Count;
end;

function TSamplesCalculator.GetSrcs(Mtrl: TMaterial; Idx: Integer): TQualitativeValue;
begin
  Result := FSources[Mtrl][Idx];
end;

function TSamplesCalculator.GetSrcField(Mtrl: TMaterial; Idx: Integer; const Index: Integer): Double;
begin
  Result := 0.;
  case Index of
    NSrcValue: Result := Srcs[Mtrl, Idx].Value;
    NSrcIUML: Result := Srcs[Mtrl, Idx].IUML;
    NSrcLog10: Result := Srcs[Mtrl, Idx].Log10;
  else
    Assert(False, 'Handled code not exists!!');
  end;
end;

function TSamplesCalculator.GetSrcIdxs(Idx: Integer): Integer;
begin
  Result := Srcs[mNil, Idx].Index;
end;

function TSamplesCalculator.GetResults(AIdx: Integer): TQualitativeResult;
begin
  Result := FResults[AIdx]
end;

function TSamplesCalculator.GetResultTexts(AIdx: Integer): String;
begin
  Result := FResults[AIdx].AsString;
end;

{ TM2Calculator }

function TM2Calculator.DoCalculate: Boolean;
var
  i: Integer;
begin
  Result := False;
  try
    FResults.Clear;
    for i := 0 to Count -1 do
      case CompareValue(SrcIumls[mNil, i], 8.0) of
        GreaterThanValue:
          FResults.Add(qrIndeterminate);

        LessThanValue,
        EqualsValue  :
          case CompareValue(IumlDeltaTBAg[i], 0.35) of
            LessThanValue:
              FResults.Add(qrNegative);

            EqualsValue     ,
            GreaterThanValue:
              case CompareValue(IumlDeltaTBAgPerNils[i], 25.0) of
                LessThanValue   :
                  FResults.Add(qrNegative);

                EqualsValue     ,
                GreaterThanValue:
                  FResults.Add(qrPositive);
              end;
          end;
      end;
  except
    on E: Exception do
    begin
      CodeSite.SendError('ErMsg on TCalcM2.Calculate', [E.Message]);
      Exit;
    end;
  end;
  Result := True;
end;

{ TM3Calculator }

function TM3Calculator.DoCalculate: Boolean;
var
  i: Integer;
begin
  Result := False;
  try
    FResults.Clear;
    for i := 0 to Count -1 do
      case CompareValue(SrcIumls[mNil, i], 8.0) of
        GreaterThanValue:
          FResults.Add(qrIndeterminate);

        LessThanValue,
        EqualsValue  :
          case CompareValue(IumlDeltaTBAg[i], 0.35) of
            LessThanValue:
              FResults.Add(IfThen(IumlDeltaMtz[i] < 0.5, qrIndeterminate, qrNegative));

            EqualsValue     ,
            GreaterThanValue:
              case CompareValue(IumlDeltaTBAgPerNils[i], 25.0) of
                LessThanValue   :
                  FResults.Add(IfThen(IumlDeltaMtz[i] < 0.5, qrIndeterminate, qrNegative));

                EqualsValue     ,
                GreaterThanValue:
                  FResults.Add(qrPositive);
              end;
          end;
      end;
  except
    on E: Exception do
    begin
      CodeSite.SendError('ErMsg on TCalcM3.Calculate', [E.Message]);
      Exit;
    end;
  end;
  Result := True;
end;

{ TMtrlCalculator }

procedure TMtrlCalculator.Clear;
begin
  FCount := 0;
  FM2Calc.Clear;
  FM3Calc.Clear;
end;

constructor TMtrlCalculator.Create;
begin
  FM2Calc := TM2Calculator.Create;
  FM3Calc := TM3Calculator.Create;
end;

destructor TMtrlCalculator.Destroy;
begin
  FreeAndNil(FM3Calc);
  FreeAndNil(FM2Calc);

  inherited;
end;

function TMtrlCalculator.Execute(const ASamples: TArray<TStringDynArray>; const ACurveFormula: TFormulaCalc): Boolean;
const
  SErMsg = 'Invalid parameter, TMtrlCalculator.Execute can handle a cmNil_Antigen or a cmNil_Antigen_Mitogen only.';
var
  si: Integer;
  LSample: TStringDynArray;
  cm: TCriteriaMaterial;
begin
  Result := False;
  for si := 0 to Length(ASamples) -1 do
  begin
    LSample := ASamples[si];
    cm := TCriteriaMaterial.CreateByItemCnt(Length(LSample));
    case cm of
      cmNil_Antigen:
        Result := FM2Calc.Execute(LSample, si, ACurveFormula);

      cmNil_Antigen_Mitogen:
        Result := FM3Calc.Execute(LSample, si, ACurveFormula);

      else
      begin
        CodeSite.SendError(SErMsg);
        //Assert(False, SErMsg);
        Exit;
      end;
    end;
    Inc(FCount);
  end;
end;

end.
