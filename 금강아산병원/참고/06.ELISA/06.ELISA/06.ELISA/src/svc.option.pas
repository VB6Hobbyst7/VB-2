unit svc.option;

interface

uses
  m.rawdata,

  mRegOption,

  System.SysUtils, System.Classes, RzCommon;

type
  TDataModule = TRegOption;
  TsvcOption = class(TDataModule)
    Reg: TRzRegIniFile;
    procedure DataModuleCreate(Sender: TObject);
  private
    FPath: String;
    FFmtPath: String;
    FPdfPath: String;
    FDataPath: String;
  protected
    function GetRegIniFile(var APath: String): TRzRegIniFile; override;
  public const
    SRootPath = 'SDBIOSENSOR\elisarprt';
    STBFmtExt = '.tff';
    SOpenDlgTBFmtFilter = 'Avaliable Format(*.tff, *.qff)|*.tff;*.qff|TB FERON Format (UNICODE)|*.tff|QuantiFERON Format (UNICODE)|*.qff';
    SDataExt = '.tbf';
    SSaveDlgDataFilter = 'TB FERON Data File(*.tbf)|*.tbf|QuantiFERON Data File(*.qdf)|*.qdf';
    SOpenDlgDataFilter = 'Avaliable Format(*.tbf, *.qft)|*.tbf;*.qff|'+SSaveDlgDataFilter;
    SCsvExt = '.csv';
    SCsvFilter = 'CSV Format(*.csv)|*.csv';
  public
    property i18nAssigned: Boolean index $0000 read GetBool write SetBool;
    property EulaAgreed: Boolean index $0001 read GetBool write SetBool;

    property MaterialCnt: Integer index $0100 read GetInteger write SetInteger;
    property Std3Col: Integer index $0101 read GetInteger write SetInteger;
    property Std3Row: Integer index $0102 read GetInteger write SetInteger;
    property Std2Col: Integer index $0103 read GetInteger write SetInteger;
    property Std2Row: Integer index $0104 read GetInteger write SetInteger;
    property StdDir: Integer index $0105 read GetInteger write SetInteger;
    property Mtrl2Dir: Integer index $0106 read GetInteger write SetInteger;
    property Mtrl3Dir: Integer index $0107 read GetInteger write SetInteger;

    property LoadFmtPath: String index $0200 read GetString write SetString;
    property SaveFmtPath: String index $0201 read GetString write SetString;
    property PDFPath: String index $0202 read GetString write SetString;
    property OpenDataPath: String index $0203 read GetString write SetString;
    property ExportDataPath: String index $0204 read GetString write SetString;
  end;

var
  svcOption: TsvcOption;

implementation

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}

uses
  System.IOUtils
  ;

{ TsvcOption }

procedure TsvcOption.DataModuleCreate(Sender: TObject);
begin
  FPath := TPath.Combine(TPath.GetDocumentsPath, 'SDBIOSENSOR\ELISA Report');
  FFmtPath := FPath + '\Format';
  FPdfPath := FPath + '\PDF';
  FDataPath := FPath + '\Save';
  if not TDirectory.Exists(FPath) then
  begin
    TDirectory.CreateDirectory(FPath);
    TDirectory.CreateDirectory(FFmtPath);
    TDirectory.CreateDirectory(FPdfPath);
  end;

  Add('common', [
    ['i18nAssigned', 'False'],
    ['EulaAgree', 'False']
  ]);

  Add('Std', [
    ['MaterialCnt', '3'],
    ['Std3Col', '3'],
    ['Std3Row', '0'],
    ['Std2Col', '5'],
    ['Std2Row', '0'],

    ['StdDir', '0'],
    ['Mtrl2Dir', '0'],
    ['Mtrl3Dir', '0']
  ]);

  Add('Path', [
    ['LoadFmt', FFmtPath],
    ['SaveFmt', FFmtPath],
    ['PDF', FPdfPath],
    ['OpenData', FDataPath],
    ['ExportData', FDataPath]
  ]);
end;

function TsvcOption.GetRegIniFile(var APath: String): TRzRegIniFile;
begin
  APath := SRootPath + '\option';
  Result := Reg;
end;

end.
