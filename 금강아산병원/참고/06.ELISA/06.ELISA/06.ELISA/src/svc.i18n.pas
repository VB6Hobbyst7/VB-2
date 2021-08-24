unit svc.i18n;

interface

uses
  m.i18n,

  mRegOption,

  System.SysUtils, System.Classes, i18nCore, System.ImageList, Vcl.ImgList, i18nCtrls, i18nLocalizer, RzCommon;

type
  TDataModule = TRegOption;
  TsvcI18n = class(TDataModule)
    Localizer: TLocalizer;
    FlagImg: TFlagImageList;
    Reg: TRzRegIniFile;
    Translator: TTranslator;
    procedure DataModuleCreate(Sender: TObject);
  private
    FFmtY, FFmtM, FFmtD: String;
    FDateSeparator: String;
    FDecimalSeparator: char;
    FTailDateSeparator: String;
    procedure AssignDateFmtSettings;
    function FmtToFmtString(const AFmt: Integer): String; overload;

    function GetDateFmts: TArray<String>;
    function GetDateFmt: String;
    function GetDateFmtString: String;
    function GetCultureInfo: TCultureInfo;
    procedure SetCultureInfo(const Value: TCultureInfo);
    function GetFmtedDates: TArray<String>;
  protected
    function GetRegIniFile(var APath: String): TRzRegIniFile; override;
  public
    function FmtToFmtString(const AFmt: TDateFmt): String; overload;
    function FmtToString(const AFmt: TDateFmt): String;
    function StdResult(const AValid: Boolean): String;

    property Culture: TCultureInfo read GetCultureInfo write SetCultureInfo;
    property CultureIdx: Integer index $0000 read GetInteger write SetInteger;

    property FmtedDates: TArray<String> read GetFmtedDates;
    property DateFmts: TArray<String> read GetDateFmts;
    property DateFmtIdx: Integer index $0001 read GetInteger write SetInteger;
    property DateFmt: String read GetDateFmt;
    property DateFmtString: String read GetDateFmtString;
    property DecimalSeparator: char read FDecimalSeparator;
  end;

var
  svcI18n: TsvcI18n;

implementation

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}

uses
  svc, svc.option,

  CodeSiteLogging, mCodeSiteHelper, System.Character, System.Math, System.StrUtils, System.DateUtils, mDateTimeHelper
  ;

const
  SYear: String = 'Year';
  SMonth: String = 'Month';
  SDay: String = 'Day';

{ TsvcI18n }

procedure TsvcI18n.AssignDateFmtSettings;
var
  LFmt: String;
  function TryAssignDateSeperator(const ASrc: String; var ASerpator: String): Boolean;
  var
    i, LPos, LLen: Integer;  
  begin
    if ASrc.Length = 0 then
      Exit(False);
      
    LPos := -1; LLen := 0; i := 0;      
    while i < LFmt.Length do
    begin
      if not CharInSet(LFmt.Chars[i], ['y', 'M', 'd']) then // do not localize
      begin
        LPos := IfThen(LPos = -1, i, LPos);
        if LLen = 0 then
          Inc(LLen)
        else if LFmt.Chars[LPos] <> LFmt.Chars[i] then
          Inc(LLen);
      end
      else if LPos > 0 then
        Break;
      Inc(i);
    end;
    Result := LPos > -1;
    if Result then
      ASerpator := ASrc.Substring(LPos, LLen);
  end;
  procedure AssingYMDFmtString(const ASrc: String; var Y, M, D: String);
  var
    LChar: Char;
  begin
    Y := ''; M := ''; D := '';
    for LChar in ASrc do
      case LChar of
        'y': Y := Y + LChar;
        'M': M := M + LChar;
        'd': D := D + LChar;
      end;
  end;
begin
  LFmt := FormatSettings.ShortDateFormat;
  if TryAssignDateSeperator(LFmt, FDateSeparator) then
  begin
    AssingYMDFmtString(LFmt, FFmtY, FFmtM, FFmtD);
    FTailDateSeparator := IfThen(LFmt.EndsWith(FDateSeparator), FDateSeparator);
  end;
end;

procedure TsvcI18n.DataModuleCreate(Sender: TObject);
begin
  FDecimalSeparator := Formatsettings.DecimalSeparator;
  Add('Common', [
    ['CultureIdx', '-1'],
    ['DateFmtIdx', TDateFmt.SystemDefault.ToInteger.ToString]
  ]);

  if CultureIdx > -1 then
    Localizer.CultureIndex := CultureIdx;
  AssignDateFmtSettings;  
end;

function TsvcI18n.FmtToFmtString(const AFmt: Integer): String;
begin
  Assert(InRange(AFmt, Low(TDateFmt).ToInteger, High(TDateFmt).ToInteger), 'Invalida Parameter. It must in rainge between Low to High of TDateFmt');
  Result := FmtToFmtString(TDateFmt.Create(AFmt));
end;

function TsvcI18n.FmtToFmtString(const AFmt: TDateFmt): String;
begin
  case AFmt of
    dfYYYYMMDD: Result := FFmtY + FDateSeparator + FFmtM + FDateSeparator + FFmtD + FTailDateSeparator;
    dfMMDDYYYY: Result := FFmtM + FDateSeparator + FFmtD + FDateSeparator + FFmtY + FTailDateSeparator;
    dfDDMMYYYY: Result := FFmtD + FDateSeparator + FFmtM + FDateSeparator + FFmtY + FTailDateSeparator;
  end;
end;

function TsvcI18n.FmtToString(const AFmt: TDateFmt): String;
begin
  case AFmt of
    dfYYYYMMDD: Result := Translator.GetText(SYear) + FDateSeparator + Translator.GetText(SMonth) + FDateSeparator + Translator.GetText(SDay) + FTailDateSeparator;
    dfMMDDYYYY: Result := Translator.GetText(SMonth) + FDateSeparator + Translator.GetText(SDay) + FDateSeparator + Translator.GetText(SYear) + FTailDateSeparator;
    dfDDMMYYYY: Result := Translator.GetText(SDay) + FDateSeparator + Translator.GetText(SMonth) + FDateSeparator + Translator.GetText(SYear) + FTailDateSeparator;
  end;
end;

function TsvcI18n.GetCultureInfo: TCultureInfo;
begin
  Result := Localizer.Culture;
end;

function TsvcI18n.GetDateFmts: TArray<String>;
var
  LFmt: TDateFmt;
begin
  Result := [];
  for LFmt := Low(TDateFmt) to High(TDateFmt) do
    Result := Result + [FmtToString(LFmt)];
end;

function TsvcI18n.GetDateFmt: String;
begin
  Result := TDateFmt.Create(DateFmtIdx).AsFmt;
end;

function TsvcI18n.GetDateFmtString: String;
begin
  Result := FmtToFmtString(DateFmtIdx);
end;

function TsvcI18n.GetFmtedDates: TArray<String>;
var
  LDate: TDateTime;
  LFmt: TDateFmt;
begin
  LDate := Now;
  Result := [];
  for LFmt := Low(TDateFmt) to High(TDateFmt) do
    Result := Result + [FmtToString(LFmt) + ' (' + LDate.ToString(FmtToFmtString(LFmt)) + ')'];
end;

function TsvcI18n.GetRegIniFile(var APath: String): TRzRegIniFile;
begin
  APath := TsvcOption.SRootPath + '\i18n';
  Result := Reg;
end;

procedure TsvcI18n.SetCultureInfo(const Value: TCultureInfo);
begin
  if Localizer.Culture <> Value then
  begin
    Localizer.Culture := Value;
    CultureIdx := Localizer.CultureIndex;

    Formatsettings.DecimalSeparator := FDecimalSeparator;
    AssignDateFmtSettings;
  end;
end;

function TsvcI18n.StdResult(const AValid: Boolean): String;
begin
  if AValid then
    Result := 'Valid ELISA test run.'
  else
    Result := '<font color="#E90309"><B>Warnning:</B> QC Criteria not met. Refer to the Package Iseert; "Quality Control of Test". ELISA is invalid and must be repeated.</font>';
end;

end.
