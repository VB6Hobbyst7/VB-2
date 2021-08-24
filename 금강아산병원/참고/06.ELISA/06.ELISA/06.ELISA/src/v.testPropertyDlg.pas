unit v.testPropertyDlg;

interface

uses
  m.test,

  mvw.vForm,

  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.WinXCtrls, i18nCore, Vcl.StdCtrls, i18nCtrls,
  Vcl.WinXCalendars, Vcl.Mask, RzEdit, Vcl.Imaging.pngimage, i18nLocalizer, Vcl.ComCtrls;

type
  TvTestPropertyDlg = class(TvDlg)
    ButtonCreate: TButton;
    Button2: TButton;
    Panel1: TPanel;
    Label1: TLabel;
    Image1: TImage;
    Label5: TLabel;
    EditTestNum: TLabeledEdit;
    EditOperator: TLabeledEdit;
    EditBatchNum: TLabeledEdit;
    Translator: TTranslator;
    DatePicker: TCalendarPicker;
    procedure FormCreate(Sender: TObject);

    procedure ButtonCreateClick(Sender: TObject);
    procedure OnEditKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure OnEditChange(Sender: TObject);private class var
  private
    function ValidateUI: Boolean;
  public
    class function Open: Boolean;
  end;

implementation

{$R *.dfm}

uses
  svc, m.rawdata,

  mEditHelper
  ;

procedure TvTestPropertyDlg.ButtonCreateClick(Sender: TObject);
var
  LProp: TGeneralInfo;
begin
  LProp.RunDate := DatePicker.Date;
  LProp.RunNumber := EditTestNum.Text;
  LProp.KitBatchNumber := EditBatchNum.Text;
  LProp.&Operator := EditOperator.Text;
  dataContainer.AssignProperties(LProp);
end;

procedure TvTestPropertyDlg.OnEditKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = VK_RETURN then
  begin
    Application.ProcessMessages;
    if ValidateUI then
      ButtonCreate.Click;
    Key := 0;
  end;
end;

procedure TvTestPropertyDlg.FormCreate(Sender: TObject);
begin
  DatePicker.DateFormat := FormatSettings.ShortDateFormat;
  DatePicker.Date := Now;
end;

procedure TvTestPropertyDlg.OnEditChange(Sender: TObject);
begin
  TThread.Queue(nil, procedure begin ValidateUI; end);
end;

class function TvTestPropertyDlg.Open: Boolean;
var
  LForm: TvTestPropertyDlg;
begin
  LForm := TvTestPropertyDlg.Create(nil);
  try
    Result := LForm.ShowModal = mrOk;
  finally
    FreeAndNil(LForm);
  end;
end;

function TvTestPropertyDlg.ValidateUI: Boolean;
var
  LEnable: Boolean;
begin
  LEnable := (EditTestNum.Text <> '') and (EditBatchNum.Text <> '')  and (EditOperator.Text <> '');
  ButtonCreate.Enabled := LEnable;
  Result := ButtonCreate.Enabled;
end;

end.
