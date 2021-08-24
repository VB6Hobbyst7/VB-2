unit v.testProperty;

interface

uses
  mvw.vForm,

  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Buttons, PngSpeedButton, Vcl.ExtCtrls, Vcl.StdCtrls, Vcl.WinXCalendars,
  Vcl.ComCtrls, i18nCore, i18nLocalizer;

type
  TvTestProperty = class(TvForm)
    PanelTitle: TPanel;
    ButtonClose: TPngSpeedButton;
    Label1: TLabel;
    EditTestNum: TLabeledEdit;
    EditBatchNum: TLabeledEdit;
    EditOperator: TLabeledEdit;
    ButtonSave: TButton;
    DatePicker: TCalendarPicker;
    Translator: TTranslator;
    procedure FormCreate(Sender: TObject);

    procedure ButtonCloseClick(Sender: TObject);
    procedure ButtonSaveClick(Sender: TObject);
    procedure DatePickerChange(Sender: TObject);
    procedure EditKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
    FOnCloseClick: TProc;
    function ValidateUI: Boolean;
    procedure DoCloseClick;
  public
    procedure Initialize;
    property OnCloseClick: TProc read FOnCloseClick write FOnCloseClick;
  end;

var
  vTestProperty: TvTestProperty;

implementation

{$R *.dfm}

uses
  svc, m.rawdata,

  System.StrUtils, System.UITypes, System.Math
  ;

{ TvTestInfo }

procedure TvTestProperty.ButtonCloseClick(Sender: TObject);
begin
  DoCloseClick;
end;

procedure TvTestProperty.ButtonSaveClick(Sender: TObject);
var
  LProp: TGeneralInfo;
begin
  LProp.RunDate := DatePicker.Date;
  LProp.RunNumber := EditTestNum.Text;
  LProp.KitBatchNumber := EditBatchNum.Text;
  LProp.&Operator := EditOperator.Text;
  dataContainer.AssignProperties(LProp);
end;

procedure TvTestProperty.DatePickerChange(Sender: TObject);
begin
  TThread.Queue(nil, procedure begin ValidateUI; end);
end;

procedure TvTestProperty.FormCreate(Sender: TObject);
begin
  DatePicker.DateFormat := FormatSettings.ShortDateFormat;
  DatePicker.Date := Now;
end;

procedure TvTestProperty.Initialize;
begin
  with dataContainer.Properties do
  begin
    DatePicker.Date := IfThen(dataContainer.HasProperties and (RunDate > 0), RunDate, Now);
    EditTestNum.Text := IfThen(dataContainer.HasProperties, RunNumber);
    EditBatchNum.Text := IfThen(dataContainer.HasProperties, KitBatchNumber);
    EditOperator.Text := IfThen(dataContainer.HasProperties, &Operator);
  end
end;

function TvTestProperty.ValidateUI: Boolean;
var
  LEnable: Boolean;
begin
  LEnable := (EditTestNum.Text <> '') and (EditBatchNum.Text <> '')  and (EditOperator.Text <> '');
  ButtonSave.Enabled := LEnable;
  Result := ButtonSave.Enabled;
end;

procedure TvTestProperty.DoCloseClick;
begin
  if Assigned(FOnCloseClick) then
    FOnCloseClick;
end;

procedure TvTestProperty.EditKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = VK_RETURN then
  begin
    TThread.Queue(nil,
      procedure
      begin
        if ValidateUI then
          ButtonSave.Click;
      end);
    Key := 0;
  end;
end;

end.
