unit v.calc.std;

interface

uses
  mvw.vForm,

  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics, AdvUtil,
  Vcl.Controls, Vcl.Grids, AdvObj, BaseGrid, AdvGrid, Vcl.StdCtrls, Vcl.Mask, RzEdit, HTMLabel, VclTee.TeeGDIPlus,
  VCLTee.TeEngine, VCLTee.Series, Vcl.ExtCtrls, VCLTee.TeeProcs, VCLTee.Chart, VCLTee.TeeFunci, i18nCore, i18nLocalizer;

type
  TvCalcStd = class(TvForm)
    LabelResult: THTMLabel;
    Grid: TAdvStringGrid;
    Chart: TChart;
    SeriesStd: TPointSeries;
    SeriesFormula: TFastLineSeries;
    LinearFunc: TCustomTeeFunction;
    Panel1: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    LabelCorreCoef: TLabel;
    EditCorreCoef: TRzNumericEdit;
    EditIntercept: TRzNumericEdit;
    EditSlope: TRzNumericEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure FormDestroy(Sender: TObject);

    procedure LinearFuncCalculate(Sender: TCustomTeeFunction; const x: Double; var y: Double);
  private const
    SResults: array[False .. True] of String = ('Fail', 'Pass');
  private
    FBmp: TBitmap;
    function ExportToString: String;
    function GetCells(c, r: Integer): String;
    function GetIntercept: String;
    function GetSlope: String;
    function GetCorreCoef: String;
  public
    procedure Initialize;
    procedure Clear;
    procedure ExportToClipboard;
    procedure ExportToFile(const AFileName: String);
    function ChartAsBmp(const w, h: Integer): TBitmap;

    property Cells[c, r: Integer]: String read GetCells;
    property Intercept: String read GetIntercept;
    property Slope: String read GetSlope;
    property CorreCoef: String read GetCorreCoef;
  end;

var
  vCalcStd: TvCalcStd;

implementation

{$R *.dfm}

uses
  m.rawdata,
  svc,

  System.UITypes, System.StrUtils, Vcl.Clipbrd, mUtils.Windows, System.DateUtils, mDateTimeHelper, System.Math,
  CodeSiteLogging
  ;

function TvCalcStd.ChartAsBmp(const w, h: Integer): TBitmap;
begin
  FBmp.SetSize(w, h);
  Chart.Draw(FBmp.Canvas, Rect(0,0, w, h));
  Result := FBmp;
end;

procedure TvCalcStd.Clear;
var
  c: Integer;
  r: Integer;
begin
  Grid.BeginUpdate;;
  try
    for c := 1 to Grid.ColCount -1 do
      for r := 1 to Grid.ColCount -1 do
        Grid.Cells[c, r] := '';

  finally
    Grid.EndUpdate;;
  end;
  SeriesStd.Clear;
  LabelResult.HTMLText.Clear;
  EditIntercept.Text := '';
  EditSlope.Text := '';
  EditCorreCoef.Text := '';
end;

procedure TvCalcStd.ExportToClipboard;
begin
  Clipboard.Clear;
  Clipboard.AsText := ExportToString;
end;

procedure TvCalcStd.ExportToFile(const AFileName: String);
var
  LBuf: TStringStream;
begin
  LBuf := TStringStream.Create(ExportToString, TEncoding.Unicode);
  try
    LBuf.SaveToFile(AFileName);
  finally
    FreeAndNil(LBuf);
  end;
end;

function TvCalcStd.ExportToString: String;
var
  LBuf: TStringWriter;
  c, r: Integer;
begin
  LBuf := TStringWriter.Create;
  try
    LBuf.WriteLine('Version: ' + ExeVersion);
    LBuf.WriteLine('Operator: ' + dataContainer.Properties.&Operator);
    LBuf.WriteLine('Kit Batch Number: ' + dataContainer.Properties.KitBatchNumber);
    LBuf.Write('Run Number' +#9);
    LBuf.Write('Run Date' +#9);
    LBuf.Write('Valid Test' +#9);
    for c := 0 to dataContainer.ColCount -1 do
      LBuf.Write(IfThen(c = 0, 'Raw Data (OD)' +#9, #9));
    LBuf.Write(#9);
    LBuf.Write('Std' +#9);
    for c := 0 to dataContainer.StdCount -1 do
      LBuf.Write(IfThen(c = 0, 'Raw Data (OD)' +#9, #9));
    LBuf.Write('Mean' +#9);
    LBuf.Write('% CV' +#9);
    LBuf.Write('QC Result' +#9);
    LBuf.Write('Correlation Coefficient');
    LBuf.WriteLine;
    for r := 0 to NRowCnt -1  do
    begin
      LBuf.Write(dataContainer.Properties.RunNumber+#9);
      LBuf.Write(dataContainer.Properties.AsRunDateStr +#9);
      LBuf.Write(IfThen(stdCalc.Valid, 'Yes', 'No')+#9);
      for c := 0 to dataContainer.ColCount -1 do
        LBuf.Write(dataContainer.MatValues[c, r] + #9);

      if r <= High(stdCalc.S1toS4) then
      begin
        LBuf.Write(#9);
        LBuf.Write(Grid.Cells[0, r +1] +#9);
        for c := 0 to stdCalc.SrcCnt -1 do
          LBuf.Write(stdCalc.Srcs[c, r].ToString +#9);
        LBuf.Write(stdCalc.Means[r].ToString +#9);
        LBuf.Write(stdCalc.CVs[r].ToString +#9);
        LBuf.Write(SResults[stdCalc.QcResults[r]] +#9);
        if r = stdCalc.S1 then
          LBuf.Write(stdCalc.ChartFormula.CorrelCoef.ToString);
      end;
      LBuf.WriteLine;
    end;
    Result := LBuf.ToString;
  finally
    FreeAndNil(LBuf);
  end;
end;

procedure TvCalcStd.FormCreate(Sender: TObject);
begin
  FBmp := TBitmap.Create;
end;

procedure TvCalcStd.FormDestroy(Sender: TObject);
begin
  FreeAndNil(FBmp);
end;

procedure TvCalcStd.FormResize(Sender: TObject);
begin
  Chart.Constraints.MaxHeight := Chart.Width
end;

function TvCalcStd.GetCells(c, r: Integer): String;
begin
  Result := Grid.Cells[c, r];
end;

function TvCalcStd.GetCorreCoef: String;
begin
  Result := Format('%s (%s)', [EditCorreCoef.Text, LabelCorreCoef.Caption]);
end;

function TvCalcStd.GetIntercept: String;
begin
  Result := EditIntercept.Text
end;

function TvCalcStd.GetSlope: String;
begin
  Result := EditSlope.Text
end;

procedure TvCalcStd.Initialize;
const
  clResults: array[False .. True] of TColor = (TColors.Red, TColors.Green);
  NIdxConc = 1; NIdxMean = 2; NIdxCV = 3; NIdxQcResult = 4;
var
  si: Integer;
begin
  stdCalc.Clear;
  stdCalc.Add(dataContainer.StdArray);
  LabelResult.HTMLText.Text := i18n.StdResult(stdCalc.Execute);
  EditIntercept.Value := stdCalc.ChartFormula.cX0;
  EditSlope.Value := stdCalc.ChartFormula.bX1;
  EditCorreCoef.Value := stdCalc.ChartFormula.CorrelCoef;
  EditCorreCoef.Color := clResults[ stdCalc.ChartFormula.CorrelCoef > 0.98 ];
  LabelCorreCoef.Caption := SResults[ stdCalc.ChartFormula.CorrelCoef > 0.98 ];
  LabelCorreCoef.Font.Color := clResults[ stdCalc.ChartFormula.CorrelCoef > 0.98 ];

  Grid.BeginUpdate;
  for si in stdCalc.S1toS4 do
  begin
    Grid.Floats[NIdxConc   , si +1] := stdCalc.TBFerons[si];
    Grid.Floats[NIdxMean   , si +1] := stdCalc.Means[si];
    if si < stdCalc.S3 then
      Grid.Cells[NIdxCV, si +1] := stdCalc.CVs[si].ToString(ffNumber, 5, 1, FormatSettings)
    else
      Grid.Cells[NIdxCV, si +1] := 'N/A';
    Grid.Cells[NIdxQcResult, si +1] := SResults[stdCalc.QcResults[si]];
    Grid.FontColors[4, si +1] := clResults[stdCalc.QcResults[si]];
  end;
  Grid.AutoFitColumns(False);
  //Grid.AutoSizeRows(False);
  Grid.EndUpdate;

  SeriesStd.BeginUpdate;
  SeriesStd.Clear;
  for si in stdCalc.S1toS3 do
    SeriesStd.AddXY(stdCalc.TBFeronLns[si], stdCalc.Lns[si]);
  SeriesStd.EndUpdate;

  LinearFunc.BeginUpdate;
  LinearFunc.StartX := MinValue(stdCalc.TBFeronLns);
  LinearFunc.ReCalculate;
  //LinearFunc.Calculate(Series, 0, 2);
  LinearFunc.EndUpdate;
end;

procedure TvCalcStd.LinearFuncCalculate(Sender: TCustomTeeFunction; const x: Double; var y: Double);
begin
  y := stdCalc.ChartFormula.bX1 * x + stdCalc.ChartFormula.cX0;
end;

end.
