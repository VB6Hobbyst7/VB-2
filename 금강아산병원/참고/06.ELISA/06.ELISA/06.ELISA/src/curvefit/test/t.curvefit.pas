unit t.curvefit;

interface

uses
  m.curvefit,

  DUnitX.TestFramework, System.Classes, System.SysUtils;

type

  [TestFixture]
  TestTCurveFit = class
  private
    FCalc: TFormulaCalc;
  public
    [Setup] procedure Setup;
    [TearDown] procedure TearDown;

    [Test] procedure TestExecute;
  end;

implementation

procedure TestTCurveFit.Setup;
begin
  FCalc := TFormulaCalc.Create;
end;

procedure TestTCurveFit.TearDown;
begin
  FreeAndNil(FCalc);
end;

procedure TestTCurveFit.TestExecute;
begin
  FCalc.Execute([45926,45952,45768,40723,40748,40625,16146,16135,16129], [45997,45988,46032,41206,41209,41247,17831,17841,17845 ], ftQuadraticEquation);
  Assert.AreEqual(0.0000, FCalc.aX2, 0.0001);
  Assert.AreEqual(0.9825, FCalc.bX1, 0.0001);
  Assert.AreEqual(-1756.2975, FCalc.cX0, 0.0001);

  Assert.IsTrue(FCalc.HasValue);
end;

initialization
  TDUnitX.RegisterTestFixture(TestTCurveFit);
end.
