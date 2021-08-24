unit v.viewNames;

interface

uses
  m.rawData,

  mvw.vForm, Spring.Collections, System.Generics.Collections,

  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, AdvUtil, Vcl.Grids, AdvObj, BaseGrid, AdvGrid, Vcl.StdCtrls, Vcl.ExtCtrls,
  Vcl.Imaging.pngimage, i18nCore, i18nLocalizer;

type
  TvViewNames = class(TvForm)
    Grid: TAdvStringGrid;
    ButtonCancel: TButton;
    ButtonOk: TButton;
    Panel1: TPanel;
    Image1: TImage;
    LabelDesc: TLabel;
    EditPrefix: TEdit;
    Label1: TLabel;
    Shape1: TShape;
    Shape2: TShape;
    Shape3: TShape;
    Translator1: TTranslator;
    procedure FormCreate(Sender: TObject);

    procedure EditPrefixKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GridCustomCellDraw(Sender: TObject; Canvas: TCanvas; ACol, ARow: Integer; AState: TGridDrawState; ARect: TRect; Printing: Boolean);
    procedure GridCanEditCell(Sender: TObject; ARow, ACol: Integer; var CanEdit: Boolean);
    procedure GridEditChange(Sender: TObject; ACol, ARow: Integer; Value: string);
    procedure GridEditCellDone(Sender: TObject; ACol, ARow: Integer);
    procedure ButtonOkClick(Sender: TObject);
  private
    FCellIDs: IDictionary<TMatPoint, String>;
    FDefaultH: Integer;
    FPrefix: String;
    FChanged: Boolean;

    procedure UpdateIDs;

    function DrawGridBg(const Canvas: TCanvas; const AState: TGridDrawState; var ARect: TRect; const c, r: Integer): Boolean;
    procedure InitGrid;
  protected
  public
    class function Open: Boolean;
  end;

implementation

{$R *.dfm}

uses
  svc,

  System.Math, System.UITypes, System.StrUtils
  ;

const
  SDescFmt = 'To edit the view names on the each cells, press enter or double click the cell and then edit it.'#13#10+
             'Prefix: %s';

procedure TvViewNames.ButtonOkClick(Sender: TObject);
var
  LItem: TPair<TMatPoint, String>;
begin
  if not FChanged then
    Exit;

  for LItem in FCellIDs do
    dataContainer.MatIDs[LItem.Key.c, LItem.Key.r] := LItem.Value;
end;

function TvViewNames.DrawGridBg(const Canvas: TCanvas; const AState: TGridDrawState; var ARect: TRect; const c,
  r: Integer): Boolean;
begin
  if gdFixed in AState then
    Exit(False);

  Canvas.Brush.Color := dataContainer.MatColors[c, r];
  Canvas.FillRect(ARect);

  Result := True;
end;

procedure TvViewNames.EditPrefixKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    vkReturn:
    begin
      FPrefix := EditPrefix.Text;
      Key := vkNone;
      TThread.Queue(nil, procedure begin UpdateIds; end);
    end;

    vkEscape:
    begin
      EditPrefix.Text := '';
      EditPrefix.TextHint := FPreFix;
      Key := vkNone;
    end;
  end;
end;

procedure TvViewNames.FormCreate(Sender: TObject);
begin
  Grid.SelectionRectangleColor := Grid.SelectionColor;
  FPrefix := 'ID ';
  FDefaultH := Grid.DefaultRowHeight;
  FCellIDs := TCollections.CreateDictionary<TMatPoint, String>;
  FChanged := False;

  InitGrid;
end;

procedure TvViewNames.GridCanEditCell(Sender: TObject; ARow, ACol: Integer; var CanEdit: Boolean);
begin
  CanEdit := not Grid.Cells[ACol, ARow].IsEmpty;
end;

procedure TvViewNames.GridCustomCellDraw(Sender: TObject; Canvas: TCanvas; ACol, ARow: Integer; AState: TGridDrawState;
  ARect: TRect; Printing: Boolean);
var
  c, r: Integer;
  LId: String;
  LMatPoint: TMatPoint;
begin
  c := ACol -1;
  r := ARow -1;

  if DrawGridBg(Canvas, AState, ARect, c, r) then
  begin
    Canvas.Font.Style := [fsBold];
    Canvas.Font.Color := clBlack;
    ARect.Inflate(-2, 0, 0, 0);
    LMatPoint := TMatPoint.Create(c, r);
    LId := ' ';
    if FCellIDs.ContainsKey(LMatPoint) then
      LId := FCellIDs[LMatPoint];
    Canvas.TextOut(ARect.Left, ARect.Top, LId);
  end;
end;

procedure TvViewNames.GridEditCellDone(Sender: TObject; ACol, ARow: Integer);
begin
  FChanged := True;
end;

procedure TvViewNames.GridEditChange(Sender: TObject; ACol, ARow: Integer; Value: string);
var
  LKey: TMatPoint;
  c, r: Integer;
begin
  c := ACol -1;
  r := ARow -1;
  LKey := TMatPoint.Create(c, r);
  if not Value.IsEmpty then
  begin
    FCellIDs[LKey] := Value;
    Grid.Invalidate;
  end;
end;

class function TvViewNames.Open: Boolean;
var
  LForm: TvViewNames;
begin
  LForm := TvViewNames.Create(nil);
  try
    Result := LForm.ShowModal = mrOk;
  finally
    FreeAndNil(LForm);
  end;
end;

procedure TvViewNames.UpdateIDs;
var
  LItem: TMatPoint;
begin
  LabelDesc.Caption := Format(Translator1.GetText(SDescFmt), [FPrefix]);

  Grid.BeginUpdate;
  for LItem in FCellIDs.Keys do
    FCellIDs[LItem] := FPreFix;
  Grid.EndUpdate;
  Grid.Invalidate;
end;

procedure TvViewNames.InitGrid;
var
  LItem: TMatPoint;
begin
  Grid.BeginUpdate;
  try
    Grid.ColCount := dataContainer.ColCount +1;
    FCellIDs.AddRange(dataContainer.MatIDArray);
    for LItem in FCellIDs.Keys do
      Grid.Cells[LItem.c +1, LItem.r +1] := FCellIDs[LItem];
  finally
    Grid.EndUpdate;
  end;
end;

end.
