unit v.i18n;

interface

uses
  mvw.vForm,

  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, i18nCore, Vcl.StdCtrls, Vcl.ExtCtrls, i18nCtrls, Vcl.Menus, RzCmboBx,
  i18nLocalizer;

type
  TvI18n = class(TvForm)
    Label2: TLabel;
    Label1: TLabel;
    ComboLang: TCultureBox;
    ComboDateFmt: TComboBox;
    Translator: TTranslator;
    procedure FormCreate(Sender: TObject);
    procedure FormActivate(Sender: TObject);

    procedure ComboLangSelect(Sender: TObject);
    procedure ComboDateFmtClick(Sender: TObject);
    procedure CtrlKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
    FActiveCtrl: TWinControl;
    FOnEnter: TProc;
    procedure UpdateDateFmt;
  protected
  public
    property OnEnter: TProc read FOnEnter write FOnEnter;
  end;

var
  vI18n: TvI18n;

implementation

{$R *.dfm}

uses
  svc,

  System.DateUtils, mDateTimeHelper, System.UITypes, mComboBoxHelper, mFontHelper
  ;

{ TvI18n }

procedure TvI18n.ComboDateFmtClick(Sender: TObject);
begin
  i18n.DateFmtIdx := ComboDateFmt.ItemIndex;
end;

procedure TvI18n.ComboLangSelect(Sender: TObject);
begin
  i18n.Culture := ComboLang.ItemSelected;
  UpdateDateFmt;
end;

procedure TvI18n.CtrlKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key = vkReturn) and (Assigned(FOnEnter)) then
  begin
    Key := 0;
    TThread.Queue(nil, procedure begin FOnEnter; end);
  end;
end;

procedure TvI18n.FormActivate(Sender: TObject);
var
  LCtrl: TWinControl;
begin
  if Assigned(FActiveCtrl) then
  begin
    LCtrl := FActiveCtrl;
    FActiveCtrl := nil;
    if LCtrl.CanFocus then
      LCtrl.SetFocus;
  end;
end;

procedure TvI18n.FormCreate(Sender: TObject);
begin
  FActiveCtrl := ComboLang;
  UpdateDateFmt;
end;

procedure TvI18n.UpdateDateFmt;
var
  LItem: String;
begin
  ComboDateFmt.Items.BeginUpdate;
  try
    ComboDateFmt.Items.Clear;
    for LItem in i18n.FmtedDates do
      ComboDateFmt.Items.Add(LItem);
    ComboDateFmt.ItemIndex := i18n.DateFmtIdx;
    ComboDateFmt.DropdownListAutoWidth;
    ComboDateFmt.SetDropdownCount;
  finally
    ComboDateFmt.Items.EndUpdate;
  end;
end;

end.
