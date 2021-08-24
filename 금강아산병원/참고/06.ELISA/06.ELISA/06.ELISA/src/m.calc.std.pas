unit m.calc.std;

interface

uses
  m.curvefit, m.rawdata,

  System.Classes, System.SysUtils, Spring.Collections, Spring
  ;

type
  TStdCalc = class
  public type
    TError = (
      ceS1MeanLow = $0000,
      ceS1CvHigh,
      ceS2CvHigh,
      ceS3Difference,
      ceS4Difference,
      ceS4MeanLow,
      ceCorrCofLow
    );
    TErrorSet = set of TError;
  const
    NS1MeanMin = 0.6;
    NS1S2CvMin = 0.15 * 100;
    NS3S4Differnce = 0.040;
    NS4MeanMin = 0.15;
    NCorrelCofMin = 0.98;
  type
    TStdInt = 0..3;
  const
    S1 = 0; S2 = 1; S3 = 2; S4 = 3;
    S1toS4: array[TStdInt] of Integer = ( S1, S2, S3, S4 );
    S1toS3: array[0..2] of Integer = ( S1, S2, S3 );
  private
    FFormula, FChartFormula: TFormulaCalc;
    FTBFeron: TArray<Double>;
    FTBFeronLn: TArray<Double>; // X
    FTBFeronLog10: TArray<Double>; // X
    FLn: TArray<Double>;        // Y
    FLog10: TArray<Double>;        // Y
    FSrc: IList<TArray<Double>>;
    FMean: TArray<Double>;
    FStdDev: TArray<Double>;
    FCV: TArray<Double>;
    FErCode: TErrorSet;
    FOnClear: TProc;
    FHasResult: Boolean;
    procedure Validate;

    function GetErMsg: String;
    function GetQcResult(Index: TStdInt): Boolean;
    function GetValid: Boolean;
    function GetItems(StdIdx: TStdInt; Row: Integer): Double;
    function GetSourceCount: Integer;
  public
    constructor Create;
    destructor Destroy; override;

    function Add(const AValues: TArray<Double>): Integer; overload;
    function Add(const AValues: TArray<TArray<Double>>): Integer; overload;
    function Execute: Boolean;
    procedure Clear;

    property HasResult: Boolean read FHasResult;
    property ChartFormula: TFormulaCalc read FChartFormula;
    property Formula: TFormulaCalc read FFormula;

    property SrcCnt: Integer read GetSourceCount;
    property Srcs[StdIdx: TStdInt; Row: Integer]: Double read GetItems; default;

    property TBFerons: TArray<Double> read FTBFeron;
    property Means: TArray<Double> read FMean;
    property CVs: TArray<Double> read FCV;
    property TBFeronLns: TArray<Double> read FTBFeronLn; // X
    property Lns: TArray<Double> read FLn;               // Y
    property StdDevs: TArray<Double> read FStdDev;
    property QcResults[Index: TStdInt]: Boolean read GetQcResult;
    property Valid: Boolean read GetValid;

    property ErMsg: String read GetErMsg;
    property ErCode: TErrorSet read FErCode;

    property OnClear: TProc read FOnClear write FOnClear;
  end;

implementation

uses
 System.Math, CodeSiteLogging, mCodeSiteHelper, System.Types, System.StrUtils, Spring.SystemUtils
 ;

{ TStdCalc }

function TStdCalc.Add(const AValues: TArray<Double>): Integer;
begin
  Result := FSrc.Add(AValues);
end;

function TStdCalc.Add(const AValues: TArray<TArray<Double>>): Integer;
var
  LArray: TArray<Double>;
begin
  for LArray in AValues do
    Add(LArray);
  Result := FSrc.Count -1;
end;

procedure TStdCalc.Clear;
begin
  FChartFormula.Initialize;
  FFormula.Initialize;
  FSrc.Clear;

  FMean := [];
  FLn := [];
  FLog10 := [];
  FStdDev := [];
  FCV := [];
  FErCode := [ceS1MeanLow, ceS1CvHigh, ceS2CvHigh, ceS3Difference, ceS4Difference, ceS4MeanLow, ceCorrCofLow];
  FHasResult := False;
  if Assigned(FOnClear) then
    FOnClear;
end;

constructor TStdCalc.Create;
begin
  FChartFormula := TFormulaCalc.Create;
  FFormula := TFormulaCalc.Create;

  FTBFeron := [4.0000, 1.0000, 0.2500];//, 0.0000];
  FTBFeronLn := [1.38629436111989, 0., -1.38629436111989];
  FTBFeronLog10 := [0.602059991327962, 0., -0.602059991327962];

  FSrc := TCollections.CreateList<TArray<Double>>;

  Clear;
end;

destructor TStdCalc.Destroy;
begin
  FreeAndNil(FChartFormula);
  FreeAndNil(FFormula);

  inherited;
end;

function TStdCalc.Execute: Boolean;
var
  si: Integer;
  function StdMaterial(const AIdx: Integer): TArray<Double>;
  var
    i: Integer;
  begin
    Result := [];
    for i := 0 to FSrc.Count -1 do
      Result := Result + [FSrc[i][AIdx]];
  end;
begin
  Result := False;
  FHasResult := True;

  FErCode := [];
  if FSrc.Count <= 1 then
    Exit;

  for si in S1toS4 do
  begin
    FMean := FMean + [Mean(StdMaterial(si))];
    FLn := FLn + [Ln(FMean[si])];
    FLog10 := FLog10 + [Log10(FMean[si])];
    FStdDev := FStdDev + [StdDev(StdMaterial(si))];
    if si < S3 then
      FCV := FCV + [FStdDev[si] / FMean[si] * 100];
  end;

  // 어떤 이유에서 인지, 참(표준)값(X축)에 FTBFeronLn을 할당한다.
  // 때문에 아래 식은 정상임.
  Validate;
  try
    FChartFormula.Execute(Copy(FLn, 0, 3), FTBFeronLn, ftLinearEquation);
    FFormula.Execute(FTBFeronLog10, Copy(FLog10, 0, 3), ftLinearEquation);
  except on E: Exception do
    begin
      CodeSite.SendError(E.Message);
      Exit;
    end;
  end;

  if FChartFormula.CorrelCoef < NCorrelCofMin then
  begin
    FErCode := FErCode + [ceCorrCofLow];
    Exit;
  end;
  Result := FErCode = [];

end;

function TStdCalc.GetErMsg: String;
begin
  Result := '';
//  case FErCode of
//    ceS1MeanLow      : Result := 'The mean value of S1 is too low. It must be greater than or equal to 0.6.';
//    ceS1S2CvHigh     : Result := 'The CV value of S2 or S3 are too high. It must be less than 15%.';
//    ceS3S4Difference : Result := 'The difference of S3 or S4 is too high. It must be ranged in 0.040.';
//    ceS4MeanLow      : Result := 'The mean value of S1 is too low. It must be greater than or equal to 0.15.';
//    ceCorrCofLow     : Result := 'The correlation coefficient value is too low. It must be greater than 0.98.';
//  end;
//  Assert(not Result.IsEmpty, 'Handled code does not exists.');
end;

function TStdCalc.GetItems(StdIdx: TStdInt; Row: Integer): Double;
begin
  Result := FSrc[StdIdx][Row];
end;

function TStdCalc.GetQcResult(Index: TStdInt): Boolean;
begin
  Result := False;
  case Index of
    S1: Result := not ((ceS1MeanLow in FErCode) or (ceS1CvHigh in FErCode));
    S2: Result := not (ceS2CvHigh in FErCode);
    S3: Result := not (ceS3Difference in FErCode);
    S4: Result := not ((ceS4Difference in FErCode) or (ceS4MeanLow in FErCode));
  end;
end;

function TStdCalc.GetSourceCount: Integer;
begin
  Result := FSrc.Count;
end;

function TStdCalc.GetValid: Boolean;
begin
  Result := FErCode = [];
end;

procedure TStdCalc.Validate;
const
  Epsilon = 0.000001;
  function MinMaxDiffCompare(const AIdx: Integer; const AComapredValue: Double): TValueRelationship;
  var
    i: Integer;
    LBuf: TArray<Double>;
  begin
    LBuf := [];
    for i := 0 to FSrc.Count -1 do
      LBuf := LBuf + [FSrc[i][AIdx]];

    Result := CompareValue(Abs(MaxValue(LBuf) - MinValue(LBuf)), AComapredValue, Epsilon);
  end;
begin
  FErCode := [];

  if CompareValue(FMean[S1], NS1MeanMin, Epsilon) = LessThanValue then
    FErCode := FErCode +[ceS1MeanLow];

  if CompareValue(FCV[S1], NS1S2CvMin, Epsilon) = GreaterThanValue then
    FErCode := FErCode +[ceS1CvHigh];

  if CompareValue(FCV[S2], NS1S2CvMin, Epsilon) = GreaterThanValue then
    FErCode := FErCode +[ceS2CvHigh];

  if MinMaxDiffCompare(S3, NS3S4Differnce) in [EqualsValue, GreaterThanValue] then
    FErCode := FErCode +[ceS3Difference];

//  if MinMaxDiffCompare(S4, NS3S4Differnce) in [EqualsValue, GreaterThanValue] then
//    FErCode := FErCode +[ceS4Difference];

  if CompareValue(FMean[S4], NS4MeanMin, Epsilon) in [GreaterThanValue] then
    FErCode := FErCode +[ceS4MeanLow];
end;

end.
