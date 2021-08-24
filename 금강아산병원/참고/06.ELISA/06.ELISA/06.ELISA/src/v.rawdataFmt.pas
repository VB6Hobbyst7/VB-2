unit v.rawdataFmt;

interface

uses
  mvw.vForm,

  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls, Vcl.StdCtrls, Vcl.Imaging.pngimage, Vcl.ExtCtrls, Vcl.Buttons,
  RzButton, RzRadChk, AdvSmoothPanel, AdvSmoothExpanderPanel, PngSpeedButton, i18nCore, i18nLocalizer;

type
  TvRawdataFmt = class(TvForm)
    PanelItr: TPanel;
    PanelM2: TPanel;
    Shape3: TShape;
    Label4: TLabel;
    PanelM2Itr: TGridPanel;
    SpeedButton4: TSpeedButton;
    SpeedButton5: TSpeedButton;
    PanelM3: TPanel;
    Shape2: TShape;
    Label3: TLabel;
    PanelM3Itr: TGridPanel;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    PanelStd: TPanel;
    Shape1: TShape;
    Label2: TLabel;
    PanelStdItr: TGridPanel;
    ButtonS1: TSpeedButton;
    ButtonS2: TSpeedButton;
    ButtonS3: TSpeedButton;
    ButtonS4: TSpeedButton;
    Panel1: TPanel;
    ButtonClose: TPngSpeedButton;
    Translator: TTranslator;
    Panel2: TPanel;
    Label1: TLabel;
    RadioSample: TRzRadioButton;
    RadioStd: TRzRadioButton;
    ButtonClear: TButton;
    PanelMaterial: TPanel;
    Label5: TLabel;
    Image7: TImage;
    RadioM2: TRzRadioButton;
    RadioM3: TRzRadioButton;
    Image6: TImage;
    PanelOrientation: TPanel;
    LabelOrientation: TLabel;
    RadioVertical: TRzRadioButton;
    RadioHorizontal: TRzRadioButton;
    RadioRandom: TRzRadioButton;
    Image1: TImage;
    Image2: TImage;
    Image5: TImage;
    procedure FormCreate(Sender: TObject);

    procedure RadioClick(Sender: TObject);
    procedure RadioItrClick(Sender: TObject);
    procedure ButtonClearClick(Sender: TObject);
    procedure ButtonCloseClick(Sender: TObject);
  private
    FOnClearSamples: TProc;
    FOnClearStds: TProc;
    FItrPanel: TPanel;
    FOnCloseClick: TProc;
    procedure UpdateItrPanels;
  public
    procedure AssignMaterial(const ACnt: Integer = 1);

    property OnCloseClick: TProc read FOnCloseClick write FOnCloseClick;
    property OnClearSamples: TProc read FOnClearSamples write FOnClearSamples;
    property OnClearStds: TProc read FOnClearStds write FOnClearStds;
  end;

var
  vRawdataFmt: TvRawdataFmt;

implementation

{$R *.dfm}

uses
  svc,
  m.rawdata,

  System.Math, System.StrUtils
  ;

procedure TvRawdataFmt.AssignMaterial(const ACnt: Integer);
begin
  dataFmter.StepIt;
end;

procedure TvRawdataFmt.ButtonClearClick(Sender: TObject);
begin
  dataFmter.Initialize;
  if RadioStd.Checked and Assigned(FOnClearStds) then
    FOnClearStds
  else if Assigned(FOnClearSamples) then
    FOnClearSamples;
end;

procedure TvRawdataFmt.ButtonCloseClick(Sender: TObject);
begin
  dataFmter.Enabled := False;
  if Assigned(FOnCloseClick) then
    FOnCloseClick;
end;

procedure TvRawdataFmt.FormCreate(Sender: TObject);
begin
  dataFmter.OnChange := procedure(AStep: Integer)
    begin
      EnumControls<TSpeedButton>(FItrPanel, procedure(ABtn: TSpeedButton)
        begin
          ABtn.Enabled := (dataFmter.Direction <> mdRandom) or (ABtn.Tag = AStep);
        end);
      ButtonClose.Enabled := dataFmter.CanInit;
      RadioStd.Enabled := dataFmter.CanInit;
      RadioSample.Enabled := dataFmter.CanInit;
      RadioM2.Enabled := dataFmter.CanInit;
      RadioM3.Enabled := dataFmter.CanInit;
      RadioVertical.Enabled := dataFmter.CanInit;
      RadioHorizontal.Enabled := dataFmter.CanInit;
    end;
  FItrPanel := PanelStd;
end;

procedure TvRawdataFmt.RadioClick(Sender: TObject);
var
  LMtrl: String;
begin
  PanelMaterial.Visible := RadioSample.Checked;
  PanelStd.Visible := RadioStd.Checked;
  PanelM3.Visible := RadioSample.Checked and RadioM3.Checked;
  PanelM2.Visible := RadioSample.Checked and RadioM2.Checked;
  LMtrl := IfThen(RadioStd.Checked, Translator.GetText('Standards'), Translator.GetText('Samples'));
  ButtonClear.Caption := Format(Translator.GetText('Clear All %s'), [LMtrl]);
  LabelOrientation.Caption := LMtrl + Translator.GetText(' Orientation');

  UpdateItrPanels;
end;

procedure TvRawdataFmt.RadioItrClick(Sender: TObject);
var
  LButton: TRzRadioButton absolute Sender;
begin
  if RadioStd.Checked then
    option.StdDir := LButton.Tag
  else if RadioM2.Checked then
    option.Mtrl2Dir := LButton.Tag
  else
    option.Mtrl3Dir := LButton.Tag;

  UpdateItrPanels;
end;

procedure TvRawdataFmt.UpdateItrPanels;
var
  LPanel: TPanel;
  LDir: TMaterialDirection;
  LButton: TRzRadioButton;
begin
  for LButton in Controls<TRzRadioButton>(PanelOrientation) do
  begin
    if RadioStd.Checked then
      LButton.Checked := LButton.Tag = option.StdDir
    else if RadioM2.Checked then
      LButton.Checked := LButton.Tag = option.Mtrl2Dir
    else
      LButton.Checked := LButton.Tag = option.Mtrl3Dir;

    if LButton.Checked then
    begin
      LDir := TMaterialDirection.Create(LButton.Tag);
      Break;
    end;
  end;

  for LPanel in Controls<TPanel>(PanelItr) do
    if LPanel.Visible then
    begin
      //ButtonClear.Top := LPanel.Height +1;
      FItrPanel := LPanel;
      EnumControls<TSpeedButton>(LPanel,
        procedure(ABtn: TSpeedButton)
        begin
          ABtn.Enabled := LDir <> mdRandom;
        end);
      dataFmter.Initialize(TCriteriaMaterial.Create(LPanel.Tag), LDir);
      Break;
    end;
end;

end.
