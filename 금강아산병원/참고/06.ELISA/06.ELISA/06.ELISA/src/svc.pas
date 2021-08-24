unit svc;

interface

uses
  svc.option, svc.i18n, m.rawdata, m.calc.std, m.calc.mtrl
  ;

function option: TsvcOption;
function i18n: TsvcI18n;
function dataContainer: TRawdataContainer;
function dataFmter: TRawdataFormater;
function stdCalc: TStdCalc;
function mtrlCalc: TMtrlCalculator;

implementation

uses
  Spring.Container
  ;

function option: TsvcOption;
begin
  Result := svcOption
end;

function i18n: TsvcI18n;
begin
  Result := svcI18n
end;

function dataContainer: TRawdataContainer;
begin
  Result := GlobalContainer.Resolve<TRawdataContainer>;
end;

function dataFmter: TRawdataFormater;
begin
  Result := GlobalContainer.Resolve<TRawdataFormater>;
end;

function stdCalc: TStdCalc;
begin
  Result := GlobalContainer.Resolve<TStdCalc>;
end;

function mtrlCalc: TMtrlCalculator;
begin
  Result := GlobalContainer.Resolve<TMtrlCalculator>;
end;

initialization
  GlobalContainer.RegisterInstance<TRawdataContainer>(TRawdataContainer.Create);
  GlobalContainer.RegisterInstance<TRawdataFormater>(TRawdataFormater.Create);
  GlobalContainer.RegisterInstance<TStdCalc>(TStdCalc.Create);
  GlobalContainer.RegisterInstance<TMtrlCalculator>(TMtrlCalculator.Create);
  GlobalContainer.Build;

finalization
  //GlobalContainer.Release(fileFmt);

end.
