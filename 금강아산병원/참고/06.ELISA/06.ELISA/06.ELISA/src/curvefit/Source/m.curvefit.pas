unit m.curvefit;

interface

uses
  mCodeSite,
  System.Classes, System.SysUtils, System.UITypes
  ;

type
  TFormulaCalc = class(TCodeSiteLogClass)
  public type
    TTerms = (
      ftLinearEquation = 2,
      ftQuadraticEquation = 3
    );
  private type
    TTermsHelper = record helper for TTerms
      class function Create(const AValue: Integer): TTerms; overload; static;
      class function Create(const AValue: String): TTerms; overload; static;
      function ToInteger: Integer;
      function ToString: String;
    end;
  private
    FHasValue: Boolean;
    FCoefs: TArray<Double>;
    FStd, FTarget: TArray<Double>;
    FCorrelCoef: Double;
    FTerms: TTerms;
    function GetCofs(const Index: Integer): Double;
  public const
    NLinear = 2;
    NQuadratic = 3;
    NIdxAx2 = 2;
    NIdxBx1 = 1;
    NIdxCx0 = 0;
  public
    constructor Create;

    procedure Initialize;

    procedure Build(const AStd, ATarget: TArray<Double>; const ATerms: Integer = 0); overload; deprecated 'Use Execute';
    function Execute(const AStd, ATarget: TArray<Double>; const ATerms: TTerms): Boolean; overload;

    property aX2: Double index NIdxAx2 read GetCofs;
    property bX1: Double index NIdxBx1 read GetCofs;
    property cX0: Double index NIdxCx0 read GetCofs;
    property CorrelCoef: Double read FCorrelCoef;
    property HasValue: Boolean read FHasValue;
    property Terms: TTerms read FTerms;

    property Std: TArray<Double> read FStd;
    property Target: TArray<Double> read FTarget;
  end;

  TFormulaTerms = TFormulaCalc.TTerms;

implementation

uses
  CurveFit,

  mGeneric, System.Math, Spring.SystemUtils
  ;

{ TFormulaCalc.TTermsHelper }

class function TFormulaCalc.TTermsHelper.Create(const AValue: Integer): TTerms;
begin
  Assert(AValue in [2, 3], 'TTerms.Create can only accept parameter by [0, 2, 3]');
  case AValue of
    2: Result := ftLinearEquation;
    3: Result := ftQuadraticEquation;
  else
    raise Exception.Create('TFormulaCalc.TTermsHelper.Create ');
  end;
end;

class function TFormulaCalc.TTermsHelper.Create(const AValue: String): TTerms;
var
  LTerm: TFormulaCalc.TTerms;
begin
  Result := ftLinearEquation;
  for LTerm in TArray<TFormulaCalc.TTerms>.Create(ftLinearEquation, ftQuadraticEquation) do
    if LTerm.ToString = AValue then
      Exit(LTerm);
end;

function TFormulaCalc.TTermsHelper.ToInteger: Integer;
begin
  case Self of
    ftLinearEquation: Result := 2;
    ftQuadraticEquation: Result := 3;
  else
    raise Exception.Create('TFormulaCalc.TTermsHelper.ToInteger');
  end;
end;

function TFormulaCalc.TTermsHelper.ToString: String;
begin
  Result := '';
  case Self of
    ftLinearEquation: Result := 'ftLinearEquation';
    ftQuadraticEquation: Result := 'ftQuadraticEquation';
  else
    Assert(False, 'Handled code not exists!!');
  end;
end;


{ TFormulaCollection }

procedure TFormulaCalc.Build(const AStd, ATarget: TArray<Double>; const ATerms: Integer);
begin
  Execute(AStd, ATarget, TTerms.Create(ATerms));
end;

function TFormulaCalc.Execute(const AStd, ATarget: TArray<Double>; const ATerms: TTerms): Boolean;
var
  LCntOfPoint: Integer;
begin
  Result := False;
  Log.EnterMethod(Self, 'Execute');
  Log.Send('AStd: %s', [TGeneric.ToLog<Double>(AStd)]);
  Log.Send('ATarget: %s', [TGeneric.ToLog<Double>(ATarget)]);
  Log.Send('ATerms', ATerms.ToString);
  FStd := AStd;
  FTarget := ATarget;
  FHasValue := False;

  LCntOfPoint := Length(FStd);
  if LCntOfPoint = 0 then
  begin
    FCoefs := [0, 1, 0];
    FCorrelCoef := 1;
    Log.ExitMethod(Self, 'Execute', Result);
    Exit(FHasValue);
  end;
  Assert(LCntOfPoint >= 2, 'Count of point is too short');
  if LCntOfPoint >= 2 then
  begin
  	FTerms := ATerms;
  	SetLength(FCoefs, FTerms.ToInteger);
    Log.Send('LCntOfPoint: %d, FTerms: : %d', [LCntOfPoint, FTerms.ToInteger]);
    try
      PolyFit(FTarget, FStd, FCoefs, FCorrelCoef, LCntOfPoint, FTerms.ToInteger);
    except
      on E: Exception do
      begin
        if E is EMatricSingluraError then
        begin
          Log.Send(E.Message);
          FCoefs := [0, 1, 0];
          FCorrelCoef := 1;
        end
        else
          Log.SendError(E.Message);
        Exit(FHasValue);
      end;
    end;
    Log.Send('FCoefs: %s, FCorrelCof: %f', [DoubleArrayToString(FCoefs), FCorrelCoef]);
    FHasValue := True;
    Result := FHasValue;
  end;
  Log.ExitMethod(Self, 'Execute', Result);
end;

constructor TFormulaCalc.Create;
begin
  Log := TCodeSiteLoggerFactory.CreateCodeSiteLogger('TFormulaCalc', TColorRec.Lightsalmon);
end;

function TFormulaCalc.GetCofs(const Index: Integer): Double;
begin
  if Length(FCoefs) > 0 then
    Result := FCoefs[Index]
  else
    case Index of
      NIdxAx2: Result := 0;
      NIdxBx1: Result := 1;
      NIdxCx0: Result := 0;
    else
      raise Exception.Create('Logical Error');
    end;
end;

procedure TFormulaCalc.Initialize;
begin
  FTarget := [];
  FStd := [];
  FCoefs := [];
  FCorrelCoef := 0;
end;

end.
