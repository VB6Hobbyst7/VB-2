unit v.about;

interface

uses
  mvw.vForm,

  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, HTMLabel, RzLabel, Vcl.Imaging.pngimage, Vcl.ExtCtrls, i18nCore,
  i18nLocalizer;

type
  TvAbout = class(TvForm)
    Image1: TImage;
    LabelVer: TRzLabel;
    HTMLabel1: THTMLabel;
    Button1: TButton;
    Translator: TTranslator;
    procedure FormCreate(Sender: TObject);
    procedure HTMLabel1AnchorClick(Sender: TObject; Anchor: string);
  private
  public
    class procedure Open;
  end;

implementation

{$R *.dfm}

uses
  v.eula,

  mUtils.Windows, StrUtils
  ;

{ TvAbout }

procedure TvAbout.FormCreate(Sender: TObject);
var
  LBuf: TArray<string>;
  LBufLen, i: Integer;
  LTail, LVer: string;
  LTailBuf: TStringList;
begin
  LBuf := ExeVersion.Split(['.']);
  LBufLen := Length(LBuf);
  LTailBuf := TStringList.Create;
  try
    LTailBuf.StrictDelimiter := True;
    for i := 0 to LBufLen -1 do
      case i of
        0: LVer := LBuf[i];
        1: LVer := LVer +'.' + LBuf[i];
        2: LTailBuf.Add(Translator.GetText('Release: ') + LBuf[i]);
        3: LTailBuf.Add(Translator.GetText('Build: ') + LBuf[i]);
      end;
    LTail := '';
    if LTailBuf.Count > 1 then
      LTail := ' (' + LTailBuf.CommaText + ')';

    LabelVer.Caption := Format(LabelVer.Caption, [LVer + LTail]);
  finally
    FreeAndNil(LTailBuf);
  end;
end;

procedure TvAbout.HTMLabel1AnchorClick(Sender: TObject; Anchor: string);
begin
  if Anchor = 'EULA' then
    TvEula.Open
end;

class procedure TvAbout.Open;
var
  LForm: TvAbout;
begin
  LForm := TvAbout.Create(nil);
  try
    LForm.ShowModal;
  finally
    FreeAndNil(LForm);
  end;
end;

end.
