unit fr.i18n;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, i18nCore, Vcl.StdCtrls, Vcl.ExtCtrls, i18nCtrls, Vcl.Menus;

type
  TfrI18n = class(TFrame)
    PopupDateFmt: TPopupMenu;
    MenuItemYYYYMMDD: TMenuItem;
    MenuItemMMDDYYYY: TMenuItem;
    MenuItemDDMMYYYY: TMenuItem;
    GridPanel1: TGridPanel;
    Label1: TLabel;
    ComboLang: TCultureBox;
    Label2: TLabel;
    EditDateFmt: TButtonedEdit;
    procedure MenuItemDateFmtClick(Sender: TObject);
    procedure ComboLangSelect(Sender: TObject);
  private
    procedure UpdateDateFmt;
  public
    procedure Initialize;


  end;

implementation

{$R *.dfm}

uses
  svc,

  System.DateUtils, mDateTimeHelper
  ;

{ TfrI18n }

procedure TfrI18n.ComboLangSelect(Sender: TObject);
begin
  i18n.Culture := ComboLang.ItemSelected;
  UpdateDateFmt;
end;

procedure TfrI18n.Initialize;
begin
  UpdateDateFmt;
end;

procedure TfrI18n.MenuItemDateFmtClick(Sender: TObject);
var
  LItem: TMenuItem absolute Sender;
begin
  LItem.Checked := True;
  i18n.DateFmtIdx := LItem.Tag;
  EditDateFmt.Text := Now.ToString(i18n.DateFmtString);
end;

procedure TfrI18n.UpdateDateFmt;
var
  LDateFmts: TArray<String>;
  LItem: TMenuItem;
begin
  EditDateFmt.Text := Now.ToString(i18n.DateFmtString);
  LDateFmts := i18n.DateFmts;
  for LItem in PopupDateFmt.Items do
  begin
    LItem.Caption := LDateFmts[LItem.Tag];
    LItem.Checked := LItem.Tag = i18n.DateFmtIdx;
  end;
end;

end.
