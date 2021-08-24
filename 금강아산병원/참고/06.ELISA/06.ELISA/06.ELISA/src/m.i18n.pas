unit m.i18n;

interface

uses
  System.SysUtils, System.Classes
  ;

type
  TDateFmt = (
    dfYYYYMMDD = 0,
    dfMMDDYYYY,
    dfDDMMYYYY
  );

  TDateFmtHelper = record helper for TDateFmt
    class function Create(const AValue: Integer): TDateFmt; static;
    class function ToFormatedArray: TArray<String>; static;
    class function SystemDefault: TDateFmt; static;
    function ToString: String;
    function AsFmt: String;
    function ToInteger: Integer;
  end;

implementation

uses
  Spring.SystemUtils, mDateTimeHelper
  ;

{ TDateFmtHelper }

function TDateFmtHelper.AsFmt: String;
begin
  Result := TEnum.GetName<TDateFmt>(Self).Substring(2);
end;

class function TDateFmtHelper.Create(const AValue: Integer): TDateFmt;
begin
  Assert(TEnum.TryParse<TDateFmt>(AValue, Result));
end;

class function TDateFmtHelper.SystemDefault: TDateFmt;
begin
  case FormatSettings.ShortDateFormat.Chars[0] of
    'd': Result := dfDDMMYYYY;
    'y': Result := dfYYYYMMDD;
    else Result := dfMMDDYYYY;
  end;
end;

class function TDateFmtHelper.ToFormatedArray: TArray<String>;
var
  LFmt: TDateFmt;
begin
  Result := [];
  for LFmt := Low(TDateFmt) to High(TDateFmt) do
    Result := Result + [LFmt.ToString];
end;

function TDateFmtHelper.ToInteger: Integer;
begin
  Result := Integer(Self);
end;

function TDateFmtHelper.ToString: String;
begin
  Result := TEnum.GetName<TDateFmt>(Self);
end;

end.
