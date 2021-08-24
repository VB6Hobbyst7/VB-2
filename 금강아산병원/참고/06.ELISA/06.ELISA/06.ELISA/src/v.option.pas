unit v.option;

interface

uses
  mvw.vForm,

  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls, AdvUtil, Vcl.Grids, AdvObj, BaseGrid, AdvGrid,
  RzLstBox, Vcl.ExtCtrls, i18nCore, i18nCtrls, i18nLocalizer;

type
  TvOption = class(TvDlg)
    Panel1: TPanel;
    ButtonCancel: TButton;
    ButtonOk: TButton;
    PanelI18n: TPanel;
    Translator: TTranslator;
    procedure FormCreate(Sender: TObject);
    procedure ButtonCancelClick(Sender: TObject);
    procedure ButtonOkClick(Sender: TObject);
  private
    FCulture: TCultureInfo;
    FDateFmtIdx: Integer;
  public
    class procedure Open;
  end;

implementation

{$R *.dfm}

uses
  svc,
  v.i18n
  ;

procedure TvOption.ButtonCancelClick(Sender: TObject);
begin
  i18n.Culture := FCulture;
  i18n.DateFmtIdx := FDateFmtIdx;
end;

procedure TvOption.ButtonOkClick(Sender: TObject);
begin
  option.i18nAssigned := True;
end;

procedure TvOption.FormCreate(Sender: TObject);
begin
  FCulture := i18n.Culture;
  FDateFmtIdx := i18n.DateFmtIdx;

  vI18n := Self.PlaceOn<TvI18n>(PanelI18n);
  vI18n.OnEnter := procedure begin ButtonOk.Click end;
end;

class procedure TvOption.Open;
var
  LForm: TvOption;
begin
  LForm := TvOption.Create(nil);
  try
    LForm.ShowModal;
  finally
    FreeAndNil(LForm);
  end;
end;

end.
