unit m.elisa;

interface

uses
  System.Classes, System.SysUtils, Vcl.Graphics
  ;

type
  TDefaultMaterial = class
  public const
    NMStdLen = 4;
    clStdMaterial: array[0..NMStdLen -1] of TColor = (
      $008C8C8C, $00AAAAAA, $00C8C8C8, $00E6E6E6
    );
    NM2Len = 44;
    clM2: array[0..NM2Len -1] of TColor = (
      $0080A6DD, $00FF8C80,  $0093EDA6, $0080A6DD,  $0080B6FF, $00FF80A3,  $0098ED91, $0080B6FF,  $0080E6FF, $00FF80CF,
      $00B6ED91, $0080E6FF,  $0080FFE6, $00C980FF,  $00D5ED91, $0080FFE6,  $0080FFB6, $00AC80FF,  $00EFE891, $0080FFB6,
      $00EFCB91, $0080FF86,  $00EFAC91, $00ACFF80,  $0080FF86, $009393EB,  $00EF8E94, $00DCFF80,  $00ACFF80, $0093B5EB,
      $00EF8EBA, $00FFF280,  $00DCFF80, $0093D7EB,  $00EF8EEF, $00FFBF80,  $00FFF280, $0093EDE7,  $009C8EEF, $00FF8C80,
      $00FFBF80, $0093EDC6,  $009C8EEF, $00FF80A3
    );
    NM3Len = 28;
    clM3: array[0..NM3Len -1] of TColor = (
      $0080A6DD, $0080FFE6, $00ACFF80,  $00FFBF80, $00FF80CF, $0093B5EB,  $0093EDA6, $00EFE891, $00C980FF,
      $0093D7EB, $0098ED91, $00EFCB91,  $0080B6FF, $0080FFB6, $00DCFF80,  $00FF8C80, $00AC80FF, $0093EDE7,
      $00B6ED91, $00EFAC91, $0080E6FF,  $0080FF86, $00FFF280, $00FF80A3,  $009393EB, $0093EDC6, $00D5ED91,
      $00EF8E94
    );
    class function IsM2StdRange(const c, r: Integer): Boolean; static;
    class function IsM3StdRange(const c, r: Integer): Boolean; static;
  public
    Values: TArray<TColor>;
    procedure FromHex(const ASrc: TArray<String>); overload;
    procedure FromHex(const ASrc: String); overload;
    function ToHex: String;
    function Count: Integer;
  end;

implementation

uses
  System.Math
  ;

{ TStdMaterial }

procedure TDefaultMaterial.FromHex(const ASrc: String);
const
  NHexLen = SizeOf(Integer) * 2;
var
  i, LCnt: Integer;
begin
  LCnt := ASrc.Length div NHexLen;
  Assert((LCnt = 44) or (LCnt = 28), 'Invalid Src');
  i := 0;
  SetLength(Values, LCnt);
  while i < LCnt do
  begin
    Values[i] := TColor(Integer.Parse(ASrc.Substring(i, NHexLen)));
    Inc(i, NHexLen);
  end;
end;

class function TDefaultMaterial.IsM2StdRange(const c, r: Integer): Boolean;
begin
  Result := InRange(c, 5, 6) and InRange(r, 0, 3);
end;

class function TDefaultMaterial.IsM3StdRange(const c, r: Integer): Boolean;
begin
  Result := InRange(c, 3, 5) and InRange(r, 0, 3);
end;

function TDefaultMaterial.Count: Integer;
begin
  Result := Length(Values);
end;

function TDefaultMaterial.ToHex: String;
var
  LBuf: TStringStream;
  LItem: TColor;
begin
  LBuf := TStringStream.Create;
  try
    for LItem in Values do
      LBuf.Write(LItem, SizeOf(LItem));
    Result := LBuf.DataString;
  finally
    FreeAndNil(LBuf);
  end;
end;

procedure TDefaultMaterial.FromHex(const ASrc: TArray<String>);
var
  LItem: String;
begin
  Values := [];
  for LItem in ASrc do
    Values := Values + [TColor(Integer.Parse(LItem))]
end;

end.
