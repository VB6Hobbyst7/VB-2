unit v.rawdata;

interface

uses
  m.rawdata,

  mvw.vForm, Spring, Spring.Collections,
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, AdvUtil, Vcl.Grids, AdvObj, BaseGrid, AdvGrid, Vcl.StdCtrls, i18nCore,
  i18nLocalizer, Vcl.ExtCtrls, HTMLabel, Vcl.ComCtrls;

type
  TvRawdata = class(TvForm)
    Grid: TAdvStringGrid;
    Translator: TTranslator;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);

    procedure GridGetAlignment(Sender: TObject; ARow, ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure GridCustomCellDraw(Sender: TObject; Canvas: TCanvas; ACol, ARow: Integer; AState: TGridDrawState; ARect: TRect; Printing: Boolean);

    procedure GridCanEditCell(Sender: TObject; ARow, ACol: Integer; var CanEdit: Boolean);
    procedure GridCellValidate(Sender: TObject; ACol, ARow: Integer; var Value: string; var Valid: Boolean);
    procedure GridMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure GridMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
    procedure GridMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure GridClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure GridMouseLeave(Sender: TObject);

  private
    FEnableFmtAssign: Boolean;
    FGridBmp: TBitmap;
    FOnPropertiesClick: TProc;
    FOnGridChange: TProc;
    function DrawGridBg(const ACanvas: TCanvas; const AState: TGridDrawState; var ARect: TRect; const c, r: Integer; const ADrawSelection: Boolean = True): Boolean;
    procedure DrawCellPoint(const ACanvases: TArray<TCanvas>; ARect: TRect; const AFontStyles: TFontStyles; const AMaterial, AValue: String);
    function GetCellAsHtmls(c, r: Integer): String;

  public
    procedure WriteInTest;
    function Paste: Boolean;
    procedure ClearStds;
    procedure ClearSamples;
    procedure BuildGridBmp;

    property CellAsHtmls[c, r: Integer]: String read GetCellAsHtmls;
    property GridAsBmp: TBitmap read FGridBmp;

    property OnProperyClick: TProc read FOnPropertiesClick write FOnPropertiesClick;
    property OnGridChange: TProc read FOnGridChange write FOnGridChange;
  end;

var
  vRawdata: TvRawdata;

implementation

{$R *.dfm}

uses
  svc,

  System.Math, CodeSiteLogging, mCodeSiteHelper, Vcl.Clipbrd, System.UITypes, System.Types, mSysUtilsEx,
  System.RegularExpressions, mRegularExpressionsHelper, System.DateUtils, mDateTimeHelper, System.StrUtils
  ;

procedure TvRawdata.ClearSamples;
begin
  dataContainer.ClearSamples;
end;

procedure TvRawdata.ClearStds;
begin
  dataContainer.ClearStds;
end;

procedure TvRawdata.DrawCellPoint(const ACanvases: TArray<TCanvas>; ARect: TRect; const AFontStyles: TFontStyles; const AMaterial, AValue: String);
var
  h: Integer;
  LCanvas: TCanvas;
begin
  for LCanvas in ACanvases do
  begin
    LCanvas.Font.Style := [fsBold];
    LCanvas.Font.Color := clRed;
    ARect.Inflate(-1, 0, 0, 0);
    LCanvas.TextOut(ARect.Left, ARect.Top, AMaterial);

    h := LCanvas.TextHeight(AMaterial);
    ARect.Inflate(-1, -h, 0, 0);
    LCanvas.Font.Style := AFontStyles;
    LCanvas.Font.Color := clBlack;
    LCanvas.TextOut(ARect.Left, ARect.Top, AValue);
  end;
end;

procedure TvRawdata.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Grid.OnCustomCellDraw := nil;
end;

procedure TvRawdata.FormCreate(Sender: TObject);
var
  LRect: TRect;
begin
  FGridBmp := TBitmap.Create;
  LRect := Grid.CellRect(0, 0) + Grid.CellRect(Grid.ColCount -1, Grid.RowCount-1);
  FGridBmp.SetSize(LRect.Width, LRect.Height);

  dataContainer.OnChange := procedure
    begin
      Grid.ColCount := dataContainer.ColCount +1;
      Grid.SelectionRectangle := False;
      Grid.Invalidate;
      BuildGridBmp;
      Grid.SelectionRectangle := True;

    end;
  dataContainer.OnCellChange := procedure(c, r: Integer)
    begin
      Grid.RepaintCell(c +1, r +1);
      //AssignGridToBmp;
    end;
end;

procedure TvRawdata.FormDestroy(Sender: TObject);
begin
  FreeAndNil(FGridBmp);
end;

function TvRawdata.GetCellAsHtmls(c, r: Integer): String;
var
  LTagFmt: String;
begin
  if c = 0 then
    Result := Grid.Cells[c, r +1]
  else if c < Grid.ColCount  then
  begin
    LTagFmt := IfThen(dataContainer.MatBlockTypes[c -1, r] <> btStandard, '%s', '<b><u>%s</u></b>');
    Result := Format(LTagFmt, [dataContainer.MatValues[c -1, r]]);
  end
  else
    Result := '';
end;

procedure TvRawdata.BuildGridBmp;
var
  LRect: TRect;
begin
  LRect := Grid.CellRect(0, 0) + Grid.CellRect(Grid.ColCount -1, Grid.RowCount-1);
  FGridBmp.SetSize(LRect.Width, LRect.Height);
  FGridBmp.Canvas.CopyRect(LRect, Grid.Canvas, LRect);
end;

procedure TvRawdata.GridCanEditCell(Sender: TObject; ARow, ACol: Integer; var CanEdit: Boolean);
var
  LValid: Boolean;
  LValue: string;
begin
  if (ARow > 0) and (ACol > 0) then
  begin
    CanEdit := not dataFmter.Enabled and dataContainer.HasData;
    if CanEdit then
    begin
      LValue := Grid.Cells[ACol, ARow];
      LValid := not LValue.IsEmpty and TRegEx.IsNumber(LValue, i18n.DecimalSeparator);
      if LValid then
        dataContainer.MatValues[ACol -1, ARow -1] := LValue;
      Grid.Cells[ACol, ARow] := dataContainer.MatValues[ACol -1, ARow -1];
    end;
  end;
end;

procedure TvRawdata.GridCellValidate(Sender: TObject; ACol, ARow: Integer; var Value: string; var Valid: Boolean);
begin
  Valid := not Value.IsEmpty and TRegEx.IsNumber(Value, i18n.DecimalSeparator);
  if Valid then
    dataContainer.MatValues[ACol -1, ARow -1] := Value;
end;

procedure TvRawdata.GridClickCell(Sender: TObject; ARow, ACol: Integer);
var
  c, r: Integer;
begin
  if (ARow = 0) or (ACol = 0) then
    Exit;

  if dataContainer.HasData and dataFmter.Enabled then
  begin
    c := ACol -1;
    r := ARow -1;
    case Grid.Cursor of
      crHandPoint:
      begin
        case dataFmter.Direction of
          mdVeritical,
          mdHorizontal:
            dataContainer.AssignManual(c, r, dataFmter.Seed, dataFmter.Direction);

          mdRandom:
            dataContainer.AssignRandom(TMatpoint.Create(c, r), dataFmter.Seed);
        end;
        dataFmter.StepIt;
        if Assigned(FOnGridChange) then
          FOnGridChange;
      end;

      crDrag:
      begin
        dataContainer.RemoveFmt(c, r);
        dataFmter.Initialize;
      end;
    end;
  end;
end;

procedure TvRawdata.GridCustomCellDraw(Sender: TObject; Canvas: TCanvas; ACol, ARow: Integer; AState: TGridDrawState;
  ARect: TRect; Printing: Boolean);
var
  c, r: Integer;
  LRect: TRect;
begin
  c := ACol -1;
  r := ARow -1;

  LRect := ARect;
  if DrawGridBg(Canvas, AState, ARect, c, r) then
  begin
    DrawGridBg(FGridBmp.Canvas, AState, LRect, c, r, False);
    DrawCellPoint([Canvas, FGridBmp.Canvas], ARect, dataContainer.MatFontStyles[c, r], dataContainer.MatMaterials[c, r], dataContainer.MatValues[c, r]);
  end;
end;

procedure TvRawdata.GridGetAlignment(Sender: TObject; ARow, ACol: Integer; var HAlign: TAlignment;
  var VAlign: TVAlignment);
begin
  VAlign := TVAlignment.tvaCenter;
  if ARow = 0 then
    HALign := TAlignment.taCenter
  else
    HALign := TAlignment.taLeftJustify;
end;

procedure TvRawdata.GridMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  FEnableFmtAssign := (Button in [mbLeft]) {and (ssCtrl in Shift)}  and (Grid.Cursor = crHandPoint);
end;

procedure TvRawdata.GridMouseLeave(Sender: TObject);
begin
  FEnableFmtAssign := False;
end;

procedure TvRawdata.GridMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
var
  c, r: Integer;
begin
  if not dataContainer.HasData or not dataFmter.Enabled then
    Grid.Cursor := crDefault
  else
  begin
    Grid.MouseToCell(X, Y, c, r);
    if (c > 0) and (r > 0) then
    begin
      Dec(c);
      Dec(r);
      Grid.Cursor := dataContainer.MatCursors[c, r];
      if FEnableFmtAssign and (Grid.Cursor = crHandPoint) then
      begin
        case dataFmter.Direction of
          mdVeritical,
          mdHorizontal:
            dataContainer.AssignManual(c, r, dataFmter.Seed, dataFmter.Direction);

          mdRandom:
            dataContainer.AssignRandom(TMatpoint.Create(c, r), dataFmter.Seed);
        end;
        dataFmter.StepIt;
        if Assigned(FOnGridChange) then
          FOnGridChange;
      end;
    end;
  end;
end;

procedure TvRawdata.GridMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if FEnableFmtAssign and (Button in [mbLeft]) then
    FEnableFmtAssign := False;
  Grid.ClearSelection;
end;

function TvRawdata.DrawGridBg(const ACanvas: TCanvas; const AState: TGridDrawState; var ARect: TRect; const c, r: Integer; const ADrawSelection: Boolean): Boolean;
begin
  if gdFixed in AState  then
    Exit(False);

  ACanvas.Brush.Color := Grid.GridLineColor;
  ACanvas.FillRect(ARect);
  ARect.Inflate(0, 0, -1, -1);

  if not Grid.AutoHideSelection and ADrawSelection and (gdSelected in AState) then
  begin
    ACanvas.Brush.Color := Grid.SelectionColor;
    ACanvas.FillRect(ARect);
    ARect.Inflate(-1, -1);
  end;
  ACanvas.Brush.Color := dataContainer.MatColors[c, r];
  ACanvas.FillRect(ARect);

  Result := True;
end;

function TvRawdata.Paste: Boolean;
const
  SErFmt = 'Invalid format on clipboard. Failed to paste into grid.';
  SErCntLess = 'The clipboard source does not match the expected column or row.';
  SErCntGreater = 'The clipboard source contains more then items than the column or row. Do you wish to proceed?';
  SErColsNotMath = 'The clipboard source has different columns count. Failed to paste into grid.';
var
  LColCnt: Integer;
  LRows, LCols: TStringList;
  r: Integer;
  LSrcGreater, LSrcLess, LSrcColNotMath: Boolean;
begin
  if not Clipboard.HasFormat(CF_TEXT) then
  begin
    MessageDlg(Translator.GetText(SErFmt), mtError, [mbOK], 0);
    Exit(False);
  end;

  LSrcColNotMath := False;
  LColCnt := -1;
  LRows := TStringList.Create;
  LCols := TStringList.Create;
  try
    LRows.Text := Clipboard.AsText;
    LSrcLess := LRows.Count < NRowCnt;
    LSrcGreater := LRows.Count > NRowCnt;
    if not LSrcLess then
      for r := 0 to LRows.Count -1 do
      begin
        LCols.StrictDelimiter := True;
        LCols.Delimiter := #9;
        LCols.DelimitedText := LRows[r];
        if LColCnt = -1 then
          LColCnt := LCols.Count;

        LSrcColNotMath := LColCnt <> LCols.Count;
        LSrcLess := LCols.Count < NMinColCnt;
        if LSrcLess or LSrcColNotMath then
          Break;
        if not LSrcGreater then
          LSrcGreater := LCols.Count > NColCnt;
      end;

    if LSrcLess then
    begin
      MessageDlg(Translator.GetText(SErCntLess), mtError, [mbOK], 0);
      Exit(False);
    end;
    if LSrcColNotMath then
    begin
      MessageDlg(Translator.GetText(SErColsNotMath), mtError, [mbOK], 0);
      Exit(False);
    end;
    if LSrcGreater then
      if MessageDlg(Translator.GetText(SErCntGreater), mtWarning, [mbYes, mbNo], 0) = mrNo then
        Exit(False);

    dataContainer.Paste(LRows, LColCnt);
    Grid.ColCount := dataContainer.ColCount +1;
    Result := True;
  finally
    FreeAndNil(LRows);
    FreeAndNil(LCols);
  end;
end;

procedure TvRawdata.WriteInTest;
const
  NCol = 12;
  NRow = 8;
var
  LColCnt: Integer;
  LRows, LCols: TStringList;
  c, r: Integer;
  LSrc: string;
begin
  for r := 0 to NRow -1 do
  begin
    for c := 0 to NCol -1 do
      LSrc := LSrc +#9;
    LSrc := LSrc +#13#10;
  end;

  LColCnt := -1;
  LRows := TStringList.Create;
  LCols := TStringList.Create;
  try
    LRows.Text := LSrc;
    for r := 0 to LRows.Count -1 do
    begin
      LCols.StrictDelimiter := True;
      LCols.Delimiter := #9;
      LCols.DelimitedText := LRows[r];
      if LColCnt = -1 then
        LColCnt := LCols.Count;
    end;
    dataContainer.Paste(LRows, LColCnt);
    Grid.ColCount := dataContainer.ColCount +1;
  finally
    FreeAndNil(LRows);
    FreeAndNil(LCols);
  end;
end;

end.
