unit v.i18nDlg;

interface

uses
  mvw.vForm,

  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, RzLabel, Vcl.ExtCtrls, Vcl.WinXCtrls, i18nCore, i18nCtrls,
  i18nLocalizer, Vcl.Menus;

type
  TvI18nDlg = class(TvForm)
    PanelContents: TRelativePanel;
    LabelI18nDesc: TLabel;
    Translator: TTranslator;
    PopupDateFmt: TPopupMenu;
    MenuItemYYYYMMDD: TMenuItem;
    MenuItemMMDDYYYY: TMenuItem;
    MenuItemDDMMYYYY: TMenuItem;
    PanelLogo: TPanel;
    LabelModuleName: TLabel;
    PanelI18n: TPanel;
    PanelButton: TPanel;
    ButtonOk: TButton;
    procedure FormCreate(Sender: TObject);

    procedure PanelContentsResize(Sender: TObject);
    procedure ButtonOkClick(Sender: TObject);
  private
  public
    class procedure Open;
  end;

var
  vI18nDlg: TvI18nDlg;

implementation

{$R *.dfm}

uses
  svc,
  v.I18n,

  mDateTimeHelper, System.DateUtils, CodeSiteLogging, mCodeSiteHelper;

procedure TvI18nDlg.ButtonOkClick(Sender: TObject);
begin
  option.i18nAssigned := True;
end;

procedure TvI18nDlg.FormCreate(Sender: TObject);
begin
  PanelLogo.BevelOuter := bvNone;
  PanelI18n.BevelOuter := bvNone;
  PanelButton.BevelOuter := bvNone;

  vI18n := Self.PlaceOn<TvI18n>(PanelI18n);
  vI18n.OnEnter := procedure begin ButtonOk.Click end;
end;

class procedure TvI18nDlg.Open;
var
  LForm: Tvi18nDlg;
begin
  LForm := Tvi18nDlg.Create(nil);
  try
    LForm.ShowModal;
  finally
    FreeAndNil(LForm);
  end;
  Application.ProcessMessages;
end;

procedure TvI18nDlg.PanelContentsResize(Sender: TObject);
var
  LPanel: TPanel absolute Sender;
begin
  PanelLogo.Top := (LPanel.Height - PanelLogo.Height) div 3;
  PanelLogo.Width := LPanel.Width;
end;

end.
