unit v.eula;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls, RzEdit, Vcl.ExtCtrls, i18nCore, i18nLocalizer;

type
  TvEula = class(TForm)
    MemoEula: TRzRichEdit;
    Translator: TTranslator;
    GridPanel1: TGridPanel;
    ButtonAgree: TButton;
    Panel1: TPanel;
    ButtonCancel: TButton;
    ButtonOk: TButton;
    Label1: TLabel;
    procedure ButtonAgreeClick(Sender: TObject);
  private
    procedure InitAgreeUI;
    procedure InitReadUI;
  public
    class function Open: Boolean;
  end;

implementation

{$R *.dfm}

uses
  svc
  ;

{ TvEula }

class function TvEula.Open: Boolean;
var
  LForm: TvEula;
begin
  LForm := TvEula.Create(nil);
  try
    if not option.EulaAgreed then
      LForm.InitAgreeUI
    else
      LForm.InitReadUI;
    Result := LForm.ShowModal = mrOk
  finally
    FreeAndNil(LForm);
  end;
end;

procedure TvEula.ButtonAgreeClick(Sender: TObject);
begin
  option.EulaAgreed := True;
end;

procedure TvEula.InitAgreeUI;
begin
  ButtonOk.Visible := False;
end;

procedure TvEula.InitReadUI;
begin
  ButtonAgree.Visible := False;
  ButtonCancel.Visible := False;
end;

end.
