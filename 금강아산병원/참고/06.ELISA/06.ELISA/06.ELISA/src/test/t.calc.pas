unit t.calc;

interface

uses
  m.calc, m.rawdata,

  DUnitX.TestFramework;

type

  [TestFixture]
  TestTStdCalc = class
  private
    FCalc: TStdCalc;
  public
    [Setup] procedure Setup;
    [TearDown] procedure TearDown;

    [Test] procedure TestStdM2;
    [Test] procedure TestStdM3;
    [Test] procedure TestErS1MeanLow;
    [Test] procedure TestErS2CvHigh;
    [Test] procedure TestErS3Difference;
    [Test] procedure TestErS4Difference;
    // Test만들기 어려워... [Test] procedure TestErS4MeanMin;
    // Test만들기 어려워... [Test] procedure TestErCorrCofLow;
  end;

  TestTMaterialCalc = class
  private
    FCalc: TMaterialCalc;
  public
    [Setup] procedure Setup;
    [TearDown] procedure TearDown;

    [Test] procedure TestStdM3;
  end;

implementation

uses
  System.Math
  ;

procedure TestTStdCalc.Setup;
begin
  FCalc := TStdCalc.Create;
end;

procedure TestTStdCalc.TearDown;
begin
  FCalc.Free;
  FCalc := nil;
end;

procedure TestTStdCalc.TestErS1MeanLow;
begin
  FCalc.Add([0.916, 0.217, 0.052, 0.007]);
  FCalc.Add([0.283, 0.217, 0.050, 0.009]);
  FCalc.Add([0.099, 0.217, 0.051, 0.008]);
  Assert.IsFalse(FCalc.Execute);
  Assert.IsTrue(ceS1MeanLow in FCalc.ErCode, 'Invalid error code');
end;

procedure TestTStdCalc.TestErS2CvHigh;
begin
  FCalc.Add([0.916, 0.217, 0.052, 0.007]);
  FCalc.Add([0.883, 0.017, 0.050, 0.009]);
  FCalc.Add([0.899, 0.817, 0.051, 0.008]);
  Assert.IsFalse(FCalc.Execute);
  Assert.AreEqual<TStdCalc.TError>(ceS1S2CvHigh, FCalc.ErCode, 'Invalid error code');
end;

procedure TestTStdCalc.TestErS3Difference;
begin
  FCalc.Add([0.916, 0.217, 0.052, 0.007]);
  FCalc.Add([0.883, 0.217, 0.011, 0.009]);
  FCalc.Add([0.899, 0.217, 0.051, 0.008]);
  Assert.IsFalse(FCalc.Execute);
  Assert.AreEqual<TStdCalc.TError>(ceS3S4Difference, FCalc.ErCode, 'Invalid error code');
end;

procedure TestTStdCalc.TestErS4Difference;
begin
  FCalc.Add([0.916, 0.217, 0.052, 0.007]);
  FCalc.Add([0.883, 0.217, 0.050, 0.059]);
  FCalc.Add([0.899, 0.217, 0.051, 0.008]);
  Assert.IsFalse(FCalc.Execute);
  Assert.AreEqual<TStdCalc.TError>(ceS3S4Difference, FCalc.ErCode, 'Invalid error code');
end;

procedure TestTStdCalc.TestStdM2;
begin
  FCalc.Add([0.916, 0.217, 0.052, 0.007]);
  FCalc.Add([0.883, 0.217, 0.050, 0.009]);
  Assert.IsTrue(FCalc.Execute);
end;

procedure TestTStdCalc.TestStdM3;
begin
  FCalc.Add([0.916, 0.217, 0.052, 0.007]);
  FCalc.Add([0.883, 0.217, 0.050, 0.009]);
  FCalc.Add([0.899, 0.217, 0.051, 0.008]);
  Assert.IsTrue(FCalc.Execute);

  Assert.AreEqual( 1.035, FCalc.Slope, 0.001, 'Slope is not match');
  Assert.AreEqual(-1.536, FCalc.Intercept, 0.001, 'Intercept is not match');
  Assert.IsTrue(FCalc.CorrelCof > 0.98, 'CorrelCof is wrong');
end;

{ TestTMaterialCalc }

procedure TestTMaterialCalc.Setup;
begin
  FCalc := TMaterialCalc.Create;
end;

procedure TestTMaterialCalc.TearDown;
begin
  FCalc.Free;
  FCalc := nil;
end;

procedure TestTMaterialCalc.TestStdM3;
begin
  FCalc.Add([0.014, 0.017, 3.005]);
  FCalc.Add([0.028, 0.028, 3.675]);
  FCalc.Add([0.013, 0.013, 3.592]);
  FCalc.Add([0.019, 0.019, 3.629]);
  FCalc.Add([0.014, 0.37 , 3.691]);
  Assert.IsTrue(FCalc.Execute);

  Assert.AreEqual(-1.853871964, FCalc.Log10s[mNil, 0], 0.0000001, 'Log10s[mNil, 0] Value is not math');
  Assert.AreEqual(-1.552841969, FCalc.Log10s[mNil, 1], 0.0000001, 'Log10s[mNil, 1] Value is not math');
  Assert.AreEqual(-1.886056648, FCalc.Log10s[mNil, 2], 0.0000001, 'Log10s[mNil, 2] Value is not math');
  Assert.AreEqual(-1.721246399, FCalc.Log10s[mNil, 3], 0.0000001, 'Log10s[mNil, 3] Value is not math');
  Assert.AreEqual(-1.853871964, FCalc.Log10s[mNil, 4], 0.0000001, 'Log10s[mNil, 4] Value is not math');
  Assert.AreEqual(-1.769551079, FCalc.Log10s[mTBAg, 0], 0.0000001, 'Log10s[mTBAg, 0] Value is not math');
  Assert.AreEqual(-1.552841969, FCalc.Log10s[mTBAg, 1], 0.0000001, 'Log10s[mTBAg, 1] Value is not math');
  Assert.AreEqual(-1.886056648, FCalc.Log10s[mTBAg, 2], 0.0000001, 'Log10s[mTBAg, 2] Value is not math');
  Assert.AreEqual(-1.721246399, FCalc.Log10s[mTBAg, 3], 0.0000001, 'Log10s[mTBAg, 3] Value is not math');
  Assert.AreEqual(-0.431798276, FCalc.Log10s[mTBAg, 4], 0.0000001, 'Log10s[mTBAg, 4] Value is not math');
  Assert.AreEqual(0.477844476, FCalc.Log10s[mMitogen, 0], 0.0000001, 'Log10s[mMitozen, 0] Value is not math');
  Assert.AreEqual(0.565257343, FCalc.Log10s[mMitogen, 1], 0.0000001, 'Log10s[mMitozen, 1] Value is not math');
  Assert.AreEqual(0.555336328, FCalc.Log10s[mMitogen, 2], 0.0000001, 'Log10s[mMitozen, 2] Value is not math');
  Assert.AreEqual(0.559786968, FCalc.Log10s[mMitogen, 3], 0.0000001, 'Log10s[mMitozen, 3] Value is not math');
  Assert.AreEqual(0.567144045, FCalc.Log10s[mMitogen, 4], 0.0000001, 'Log10s[mMitozen, 4] Value is not math');


  Assert.AreEqual(0.071427, FCalc.IUML[mNil, 0], 0.00001, 'IUML[mNil, 0] Value is not math');
  Assert.AreEqual(0.139526, FCalc.IUML[mNil, 1], 0.00001, 'IUML[mNil, 1] Value is not math');
  Assert.AreEqual(0.066492, FCalc.IUML[mNil, 2], 0.00001, 'IUML[mNil, 2] Value is not math');
  Assert.AreEqual(0.095935, FCalc.IUML[mNil, 3], 0.00001, 'IUML[mNil, 3] Value is not math');
  Assert.AreEqual(0.071427, FCalc.IUML[mNil, 4], 0.00001, 'IUML[mNil, 4] Value is not math');
  Assert.AreEqual(0.086162, FCalc.IUML[mTBAg, 0], 0.00001, 'IUML[mTBAg, 0] Value is not math');
  Assert.AreEqual(0.139526, FCalc.IUML[mTBAg, 1], 0.00001, 'IUML[mTBAg, 1] Value is not math');
  Assert.AreEqual(0.066492, FCalc.IUML[mTBAg, 2], 0.00001, 'IUML[mTBAg, 2] Value is not math');
  Assert.AreEqual(0.095935, FCalc.IUML[mTBAg, 3], 0.00001, 'IUML[mTBAg, 3] Value is not math');
  Assert.AreEqual(1.688818, FCalc.IUML[mTBAg, 4], 0.00001, 'IUML[mTBAg, 4] Value is not math');
  Assert.AreEqual(12.773143, FCalc.IUML[mMitogen, 0], 0.00001, 'IUML[mMitozen, 0] Value is not math');
  Assert.AreEqual(15.514529, FCalc.IUML[mMitogen, 1], 0.00001, 'IUML[mMitozen, 1] Value is not math');
  Assert.AreEqual(15.175915, FCalc.IUML[mMitogen, 2], 0.00001, 'IUML[mMitozen, 2] Value is not math');
  Assert.AreEqual(15.326896, FCalc.IUML[mMitogen, 3], 0.00001, 'IUML[mMitozen, 3] Value is not math');
  Assert.AreEqual(15.579774, FCalc.IUML[mMitogen, 4], 0.00001, 'IUML[mMitozen, 4] Value is not math');

  Assert.AreEqual(0.01474, FCalc.DiffTBAgNil[0], 0.00001, 'DiffTBAgNil[0] Value is not math');
  Assert.AreEqual(0.00000, FCalc.DiffTBAgNil[1], 0.00001, 'DiffTBAgNil[1] Value is not math');
  Assert.AreEqual(0.00000, FCalc.DiffTBAgNil[2], 0.00001, 'DiffTBAgNil[2] Value is not math');
  Assert.AreEqual(0.00000, FCalc.DiffTBAgNil[3], 0.00001, 'DiffTBAgNil[3] Value is not math');
  Assert.AreEqual(1.61739, FCalc.DiffTBAgNil[4], 0.00001, 'DiffTBAgNil[4] Value is not math');
  Assert.AreEqual(  20.62962, FCalc.DiffTBAgNilPerNil[0], 0.00001, 'DiffTBAgNilPerNil[0] Value is not math');
  Assert.AreEqual(   0.00000, FCalc.DiffTBAgNilPerNil[1], 0.00001, 'DiffTBAgNilPerNil[1] Value is not math');
  Assert.AreEqual(   0.00000, FCalc.DiffTBAgNilPerNil[2], 0.00001, 'DiffTBAgNilPerNil[2] Value is not math');
  Assert.AreEqual(   0.00000, FCalc.DiffTBAgNilPerNil[3], 0.00001, 'DiffTBAgNilPerNil[3] Value is not math');
  Assert.AreEqual(2264.41215, FCalc.DiffTBAgNilPerNil[4], 0.00001, 'DiffTBAgNilPerNil[4] Value is not math');
  Assert.AreEqual(12.70172, FCalc.DiffMtzNil[0], 0.00001, 'DiffMtzNil[0] Value is not math');
  Assert.AreEqual(15.37500, FCalc.DiffMtzNil[1], 0.00001, 'DiffMtzNil[1] Value is not math');
  Assert.AreEqual(15.10942, FCalc.DiffMtzNil[2], 0.00001, 'DiffMtzNil[2] Value is not math');
  Assert.AreEqual(15.23096, FCalc.DiffMtzNil[3], 0.00001, 'DiffMtzNil[3] Value is not math');
  Assert.AreEqual(15.50835, FCalc.DiffMtzNil[4], 0.00001, 'DiffMtzNil[4] Value is not math');

  Assert.AreEqual<TMaterialCalc.TResult>(mrNegative, FCalc.Results[0], 'Results[0] is not math');
  Assert.AreEqual<TMaterialCalc.TResult>(mrNegative, FCalc.Results[1], 'Results[1] is not math');
  Assert.AreEqual<TMaterialCalc.TResult>(mrNegative, FCalc.Results[2], 'Results[2] is not math');
  Assert.AreEqual<TMaterialCalc.TResult>(mrNegative, FCalc.Results[3], 'Results[3] is not math');
  Assert.AreEqual<TMaterialCalc.TResult>(mrPositive, FCalc.Results[4], 'Results[4] is not math');
end;

initialization
  TDUnitX.RegisterTestFixture(TestTStdCalc);
  TDUnitX.RegisterTestFixture(TestTMaterialCalc);

end.
