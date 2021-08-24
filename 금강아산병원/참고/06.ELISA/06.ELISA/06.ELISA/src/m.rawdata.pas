unit m.rawdata;

interface

uses
  System.Classes, System.SysUtils, System.IniFiles, Spring.Collections, System.UITypes, Spring,
  System.Generics.Collections, System.Generics.Defaults, System.SyncObjs
  ;


const
  NColCnt = 12;
  NMinColCnt = 2;
  NRowCnt =  8;

type
  TMatPoint = record
    c,
    r: Integer;
    constructor Create(const ACol, ARow: Integer);
  end;
  TMatStrPair = TPair<TMatPoint, String>;

  TMat<T> = array[0..NColCnt -1, 0..NRowCnt -1] of T;

  TGeneralInfo = record
  private const
    NRunNumber      = 0;
    NDataRows       = 1;
    NDataCols       = 2;
  private
    FHasValue: Boolean;
    FCurveType: Integer;
    FRunDate: TDateTime;
    FRunNumber: String;
    FKitBatchNumber: String;
    FOperator: String;
    FDataRows: Integer;
    FDataCols: Integer;
    FDecimalSeparator: Char;
    FDateSeparator: Char;
    procedure SetInteger(const AIdx, AValue: Integer);
    procedure SetOperator(const Value: String);
    procedure SetRunDate(const Value: TDateTime);
    procedure GetKitBatchNumber(const Value: String);
    function GetAsRunDateStr: String;
    procedure SetRunNumber(const Value: String);
    function GetDateSeparator: Char;
  public
    class function Default: TGeneralInfo; static;

    constructor Create(const ARunDate: TDateTime; const ARunNumber, AKitBatchNumber, AOperator: String); overload;
    constructor Create(const ASectionValues: TStringList); overload;

    procedure Clear;
    procedure WriteFmt(const ABuf: TTextWriter; const AContainProperty: Boolean = False);
    procedure Assign(const AValue: TGeneralInfo);

    property HasValue         : Boolean                           read FHasValue;
    property CurveType        : Integer                           read FCurveType;
    property RunDate          : TDateTime                         read FRunDate          write SetRunDate;
    property AsRunDateStr     : String read GetAsRunDateStr;
    property RunNumber        : String read FRunNumber        write SetRunNumber;
    property KitBatchNumber   : String                            read FKitBatchNumber   write GetKitBatchNumber;
    property &Operator        : String                            read FOperator         write SetOperator;
    property DataRows         : Integer   index NDataRows         read FDataRows         write SetInteger;
    property DataCols         : Integer   index NDataCols         read FDataCols         write SetInteger;
    property DecimalSeparator : Char      read FDecimalSeparator write FDecimalSeparator;
    property DateSeparator    : Char      read GetDateSeparator    write FDateSeparator;
  end;

  TCellPoint = record
  private
    FIdx: Integer;
    function GetValueAssigned: Boolean;
  public
    Col,
    Row: Integer;
    Value: String;
    constructor Create(const ASrc: String); overload;
    constructor Create(const ACol, ARow, AIdx: Integer; const AValue: String); overload;

    function MatPoint: TMatPoint;

    property Idx: Integer read FIdx;
    property ValueAssigned: Boolean read GetValueAssigned;
  end;

  PCellPoint = ^TCellPoint;

  TBlockType = (
    btMaterial= 0,
    btStandard,
    btEmpty
  );
  PBlockType = ^TBlockType;

  TBlockTypeHelper = record helper for TBlockType
    class function Create(const AValue: Integer): TBlockType; static;
    function ToInteger: Integer;
    function ToString: String;
  end;

  TCriteriaMaterial = (
    cmNil_Antigen = 0,
    cmNil_Antigen_Mitogen,
//    cmNot_Specified,
//    cmGold_Nil_ESAT6_Cfp9_Mitogen,
//    cmNil_CMV_Mitogen,
//    cmNil_CMV_Mitogen_DilCMV_DilMit,
//    cmNil_DilutedCMV_DilutedMitogen,
//    cmNil_CMV_Mitogen_DilutedCMV,
//    cmNil_DilutedCMV_Mitogen,
    cmStandard,

    cmNone
  );
  TCriteriaMaterialSet = set of TCriteriaMaterial;
  TCrtrMtrlSampleRange = cmNil_Antigen .. cmNil_Antigen_Mitogen;
const
  TCrtrMtrlSampleSet: TCriteriaMaterialSet = [cmNil_Antigen, cmNil_Antigen_Mitogen];
type
  PCriteriaMaterial = ^TCriteriaMaterial;
  TCriteriaMaterialHelper = record helper for TCriteriaMaterial
  private
    function GetMaterialString(Index: Integer): String;
  public
    class function Create(const AValue: Integer): TCriteriaMaterial; static;
    class function CreateByItemCnt(const AValue: Integer): TCriteriaMaterial; static;

    function ToInteger: Integer;
    function ToString: String;
    function PointLen: Integer;
    function ToBlock: TBlockType;
    function ToItemCnt: Integer;
    property MaterialString[Index: Integer]: String read GetMaterialString;

    function DefaultIdx(const c, r: Integer): Integer;
    function DefaultPointIdx(const c, r: Integer): Integer;
    function DefaultStdRange(const c, r: Integer): Boolean;
    function DefaultPointsLen(const c, r: Integer): Integer;
    function DefaultMaterial(const c, r: Integer): TCriteriaMaterial;
    function DefaultBlock(const c, r: Integer): TBlockType;
  end;

  TMaterial = (
    mNil,
    mTBAg,
    mMitogen
  );
const
  TMaterialAll: array[0..2] of TMaterial = (mNil, mTBAg, mMitogen);
type
  TMaterialHelper = record helper for TMaterial
    class function Create(const AValue: Integer): TMaterial; static;
    class function ToArray: TArray<TMaterial>; static;
  end;

  TSectionCell = record
  public
    FAssigned: Boolean;
    FPoints: TArray<TCellPoint>;
    function GetPointLen: Integer;
    function GetIsEmpty: Boolean;
  private
    function GetMaterialText(PointIdx: Integer): String;
    function GetLastPoint: TCellPoint;
  public
    CriteriaMaterial: TCriteriaMaterial;
    FixedID: Integer;
    ID: String;
    BlockType: TBlockType;

    constructor CreateFromSection(const AIdx: Integer; ASection: TStrings; const AFree: Boolean = True); overload;
    constructor CreateEmpty(const AFixedIdx, c, r: Integer; const AValue: String); overload;
    constructor Create(const AFixedID: Integer; const ABlock: TBlockType; AMaterial: TCriteriaMaterial); overload;
    constructor Create(const AFixedID: Integer; const ABlock: TBlockType; AMaterial: TCriteriaMaterial; const APoint: TCellPoint); overload;

    procedure Write(const AIdx: Integer; const ABuf: TTextWriter);

    procedure AddPoint(AValue: TCellPoint); overload;
    function AddPoint(const c, r: Integer; AValue: String): TCellPoint; overload;
    procedure ClearPoint;
    function Value(c, r: Integer): String;
    function PointIdxOf(const c, r: Integer): Integer;
    function ToDoubleArray: TArray<Double>;
    function ToStringArray: TStringDynArray;
    procedure SetPointsByMat(MatPoint: TMatPoint; const Value: TCellPoint);

    property IsEmptyBlock: Boolean read GetIsEmpty;
    property PointLen: Integer read GetPointLen;
    property Points: TArray<TCellPoint> read FPoints write FPoints;
    property LastPoint: TCellPoint read GetLastPoint;
    property MaterialText[PointIdx: Integer]: String read GetMaterialText;
  end;

  TDefaultMaterial = class
  public const
    NMStdLen = 4;
    clMStd: array[0..NMStdLen -1] of TColor = (
      $008C8C8C, $00AAAAAA, $00C8C8C8, $00E6E6E6
    );
    // Grid에 Std Material을 할당안한다면, 4개의 셀 색상이 필요하다.
    NExtraColorLen = 4;
    NM2Len = 44;
    NM2MaxLen = NM2Len + NExtraColorLen;
    clM2: array[0..NM2MaxLen -1] of TColor = (
      $0080A6DD, $00FF8C80,  $0093EDA6, $0080A6DD,  $0080B6FF, $00FF80A3,  $0098ED91, $0080B6FF,  $0080E6FF, $00FF80CF,
      $00B6ED91, $0080E6FF,  $0080FFE6, $00C980FF,  $00D5ED91, $0080FFE6,  $0080FFB6, $00AC80FF,  $00EFE891, $0080FFB6,
      $00EFCB91, $0080FF86,  $00EFAC91, $00ACFF80,  $0080FF86, $009393EB,  $00EF8E94, $00DCFF80,  $00ACFF80, $0093B5EB,
      $00EF8EBA, $00FFF280,  $00DCFF80, $0093D7EB,  $00EF8EEF, $00FFBF80,  $00FFF280, $0093EDE7,  $009C8EEF, $00FF8C80,
      $00FFBF80, $0093EDC6,  $009C8EEF, $00FF80A3,
      //Extra color
      $00FF80CF, $00C980FF, $00AC80FF, $009393EB
    );
    NM3Len = 28;
    NM3MaxLen = NM3Len + NExtraColorLen;
    clM3: array[0..NM3MaxLen -1] of TColor = (
      $0080A6DD, $0080FFE6, $00ACFF80,  $00FFBF80, $00FF80CF, $0093B5EB,  $0093EDA6, $00EFE891, $00C980FF,
      $0093D7EB, $0098ED91, $00EFCB91,  $0080B6FF, $0080FFB6, $00DCFF80,  $00FF8C80, $00AC80FF, $0093EDE7,
      $00B6ED91, $00EFAC91, $0080E6FF,  $0080FF86, $00FFF280, $00FF80A3,  $009393EB, $0093EDC6, $00D5ED91,
      $00EF8E94,
      //Extra color
      $00EF8EBA, $00EF8EEF, $009C8EEF, $009C8EEF
    );
    clRandom: array[0..47] of TColor = (
      $0080A6DD, $0080B6FF, $0080E6FF, $0080FFE6, $0080FFB6, $0080FF86, $00ACFF80, $00DCFF80, $00FFF280, $00FFBF80,
      $00FF8C80, $00FF80A3, $00FF80CF, $00C980FF, $00AC80FF, $009393EB, $0093B5EB, $0093D7EB, $0093EDE7, $0093EDC6,
      $0093EDA6, $0098ED91, $00B6ED91, $00D5ED91, $00EFE891, $00EFCB91, $00EFAC91, $00EF8E94, $00EF8EBA, $00EF8EEF,
      $009C8EEF, $009C8EEF, $0080A6DD, $0080B6FF, $0080E6FF, $0080FFE6, $0080FFB6, $0080FF86, $00ACFF80, $00DCFF80,
      $00FFF280, $00FFBF80, $00FF8C80, $00FF80A3, $00FF80CF, $00C980FF, $00AC80FF, $009393EB
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

  TMaterialDirection = (
    mdVeritical = 0,
    mdHorizontal,
    mdRandom
  );
  TMaterialDirectionHelper = record helper for TMaterialDirection
    class function Create(const AValue: Integer): TMaterialDirection; static;
    function ToInteger: Integer;
    function ToString: String;
  end;

  TRawdataFormater = class
  private
    FDir: TMaterialDirection;
    FSeed: TCriteriaMaterial;
    FCapacity, FRemainCnt: Integer;
    FOnChange: TProc<Integer>;
    FEnabled: Boolean;
    function GetCanInit: Boolean;
    procedure SetEnabled(const Value: Boolean);
    procedure DoChange;
  public
    constructor Create;

    procedure Initialize(const ASeed: TCriteriaMaterial; const ADir: TMaterialDirection); overload;
    procedure Initialize; overload;
    procedure StepIt;

    property Enabled: Boolean read FEnabled write SetEnabled;
    property CanInit: Boolean read GetCanInit;
    property Seed: TCriteriaMaterial read FSeed;
    property Direction: TMaterialDirection read FDir;
    property Capacity: Integer read FCapacity;

    property OnChange: TProc<Integer> read FOnChange write FOnChange;
  end;

  ISectionCellDic = IDictionary<Integer, TSectionCell>;
  TMaterialTextDic = class
  private
    FDic: IDictionary<TMatPoint, String>;
    function GetItems(c, r: Integer): String;
  public
    constructor Create;

    procedure Add(const ABlock: TBlockType; const AFixedId, APointIdx: Integer; const APoint: TCellPoint);
    procedure AddOrSetValue(const ABlock: TBlockType; const AFixedId, APointIdx: Integer; const APoint: TCellPoint);
    procedure Clear;
    procedure Remove(const AKey: TMatPoint);

    property Items[c, r: Integer]: String read GetItems; default;
  end;

  TSectionCellContainer = class
  private
    function GetLast: TSectionCell;
    function GetMatMaterialText(c, r: Integer): String;
    function GetMatIDs(c, r: Integer): String;
    function GetMatValues(c, r: Integer): String;
    procedure SetMatValues(c, r: Integer; const Value: String);
    procedure SetMatIDs(c, r: Integer; const Value: String);
  protected
    FOrderedByAddedSeq: IList<Integer>;
    FIdxs: IDictionary<TMatPoint, Integer>;
    FPoints: IDictionary<TMatPoint, TCellPoint>;
    FMtrlText: TMaterialTextDic;
    FCells: IDictionary<Integer, TSectionCell>;
    FIDs: IDictionary<TMatPoint, String>;
    function GetCount: Integer;
    function GetMatPoints(c, r: Integer): TCellPoint;
    procedure SetMatPoints(c, r: Integer; const Value: TCellPoint);
    function GetSectionCells(Idx: Integer): TSectionCell;
    procedure SetSectionCells(Idx: Integer; const Value: TSectionCell);
  public
    constructor Create;
    destructor Destroy; override;

    function AddSectionCell(const ADstIdx: Integer; const AValue: TSectionCell): TSectionCell;
    function AddPoint(const ADstIdx: Integer; const AValue: TCellPoint): TSectionCell;
    procedure Clear;
    function Contains(const AIdx: Integer): Boolean;
    function Extract(const AIdx: Integer; var AValue: TSectionCell): Boolean; overload;

    function Keys: TArray<Integer>;
    function SortedByPointKeys: TArray<Integer>;

    property Count: Integer read GetCount;
    property Last: TSectionCell read GetLast;
    property MatPoints[c, r: Integer]: TCellPoint read GetMatPoints write SetMatPoints;
    property MatValues[c, r: Integer]: String read GetMatValues write SetMatValues;
    property MatMaterialText[c, r: Integer]: String read GetMatMaterialText;
    property MatIDs[c, r: Integer]: String read GetMatIDs write SetMatIDs;
    property SectionCells[Idx: Integer]: TSectionCell read GetSectionCells write SetSectionCells; default;
  end;

  TRawDataCellContainer = class(TSectionCellContainer)
  public
    function Extract(const AKey: TMatPoint; var AValue: TSectionCell): Boolean; overload;
    function FindEmptyPoint(const ADir: TMaterialDirection; var APt: TMatPoint; var AValue: TSectionCell): Boolean;
  end;

  TRawDataSeed = class(TStringList)
  private
    function GetMats(c, r: Integer): String;
    procedure SetMats(c, r: Integer; const Value: String);
  public
    function Cols(const r: Integer; var ACols: TStringList): Boolean;

    property Mats[c, r: Integer]: String read GetMats write SetMats;
  end;

  TRawdataContainer = class
  private
    FProperties: TGeneralInfo;
    FLock: TObject;
    FSeed: TRawDataSeed;
    FRawData: TRawDataCellContainer;
    FStds, FMtrls: TSectionCellContainer;
    FBlockTypes: TMat<TBlockType>;
    FCellIdx: TMat<Integer>;
    FOnChange: TProc;
    FColCount: Integer;
    FOnCellChange: TProc<Integer, Integer>;
    FOnPropertyChange: TProc;
    procedure DoChange;
    procedure BuildRawData;
    procedure BuildSeed(const ASrc: TStringList);
    procedure ClearSectionCells(const ADic: TSectionCellContainer);

    function GetHasData: Boolean;
    function GetMatColors(c, r: Integer): TColor;
    function GetMatMaterial(c, r: Integer): String;
    function GetMatValues(c, r: Integer): String;
    procedure SetMatValues(c, r: Integer; const Value: String);
    function GetMatFontStyles(c, r: Integer): TFontStyles;
    function GetMatCursors(c, r: Integer): TCursor;
    function GetMatIDs(c, r: Integer): String;
    function GetHasStd: Boolean;
    procedure SetMatIDs(c, r: Integer; const Value: String);
    function GetIDs(Idx: Integer): String;
    function GetCellCount: Integer;
    function GetStdCount: Integer;
    function GetPointsByIds(Idx: Integer): TArray<TCellPoint>;
    function GetCanCalc: Boolean;
  public
    constructor Create;
    destructor Destroy; override;

    procedure Paste(const ASrc: TStringList; const AColCount: Integer);
    procedure AssignDefault(const AMaterial: TCriteriaMaterial);
    procedure AssignManual(c, r: Integer; const AMaterial: TCriteriaMaterial; const ADir: TMaterialDirection);
    procedure AssignRandom(const APoint: TMatPoint; const AMaterial: TCriteriaMaterial);
    procedure RemoveFmt(c, r: Integer);
    procedure AssignFmtSet(const AValue: TFormatSettings);
    procedure LoadFromFile(const AFileName: String);
    procedure SaveToFile(const AFileName: String; const AContainsProperty: Boolean = False);

    procedure Clear(const AClearWithProp: Boolean = True);
    procedure ClearStds;
    procedure ClearSamples;

    procedure AssignProperties(const AProperty: TGeneralInfo);

    function MatIDArray: TArray<TMatStrPair>;
    function IDArray: TArray<String>;
    function StdArray: TArray<TArray<Double>>;
    function MtrlArray: TArray<TStringDynArray>;

    property HasProperties: Boolean read FProperties.FHasValue;
    property Properties: TGeneralInfo read FProperties;

    property HasData: Boolean read GetHasData;
    property HasStd: Boolean read GetHasStd;
    property CanCalc: Boolean read GetCanCalc;

    property ColCount: Integer read FColCount;
    property CellCount: Integer read GetCellCount;
    property StdCount: Integer read GetStdCount;
    property MatColors[c, r: Integer]: TColor read GetMatColors;
    property MatMaterials[c, r: Integer]: String read GetMatMaterial;
    property MatValues[c, r: Integer]: String read GetMatValues write SetMatValues;
    property PointsByIds[Idx: Integer]: TArray<TCellPoint> read GetPointsByIds;
    property IDs[Idx: Integer]: String read GetIDs;
    property MatIDs[c, r: Integer]: String read GetMatIDs write SetMatIDs;
    property MatFontStyles[c, r: Integer]: TFontStyles read GetMatFontStyles;
    property MatBlockTypes: TMat<TBlockType> read FBlockTypes;
    property MatCursors[c, r: Integer]: TCursor read GetMatCursors;

    property OnChange: TProc read FOnChange write FOnChange;
    property OnCellChange: TProc<Integer, Integer> read FOnCellChange write FOnCellChange;
    property OnPropertyChange: TProc read FOnPropertyChange write FOnPropertyChange;
  end;

function IfThen(const AValue: Boolean; const ATrue, AFalse: TCriteriaMaterial): TCriteriaMaterial; overload;
function IfThen(const AValue: Boolean; const ATrue, AFalse: TBlockType): TBlockType; overload;
function IfThen(const AValue: Boolean; const ATrue, AFalse: TMaterialDirection): TMaterialDirection; overload;

implementation

uses
  System.Math, mStringListHelper, Spring.SystemUtils, mDateTimeHelper, System.DateUtils, System.StrUtils,
  CodeSiteLogging, mCodeSiteHelper, mExceptions, System.Types
  ;

const
  SGeneralInfo = 'GeneralInfo';
  SCell = 'Cell';
  SCellFmt = 'Cell%d';
  SPoint = 'Point';
  SBlockType = 'BlockType';

function IfThen(const AValue: Boolean; const ATrue, AFalse: TCriteriaMaterial): TCriteriaMaterial; overload;
begin
  if AValue then
    Exit(ATrue)
  else
    Exit(AFalse);
end;

function IfThen(const AValue: Boolean; const ATrue, AFalse: TBlockType): TBlockType; overload;
begin
  if AValue then
    Exit(ATrue)
  else
    Exit(AFalse);
end;

function IfThen(const AValue: Boolean; const ATrue, AFalse: TMaterialDirection): TMaterialDirection; overload;
begin
  if AValue then
    Exit(ATrue)
  else
    Exit(AFalse);
end;

function IfThen(const ACondition: Boolean; const ATrue, AFalse: TSectionCellContainer): TSectionCellContainer; overload;
begin
  if ACondition then Exit(ATrue)
  else               Exit(AFalse);
end;

type
  TQFmtBuf = class
  private type
    TSection = class
    private
      FList: TStringList;
      FPoints: IList<TCellPoint>;
      FMaterial: TCriteriaMaterial;
      FBlock: TBlockType;
      FFixedID: Integer;
      FID: String;
        function GetPointCnt: Integer;
    public
      constructor Create(const ASrc: TStringList);
      destructor Destroy; override;

      property PointCnt: Integer read GetPointCnt;
    end;
  private
    FProperties: TGeneralInfo;
    FSecNames: TStringList;
    FSections: IDictionary<Integer, TSection>;
    FIni: TMemIniFile;
    FCellCnt: Integer;
    function GetSections(Idx: Integer): TSection;
    function GetBlockType(Idx: Integer): TBlockType;
    function GetFixedID(Idx: Integer): Integer;
    function GetMaterial(Idx: Integer): TCriteriaMaterial;
    function GetPointCnt(Idx: Integer): Integer;
    function GetPoints(i, j: Integer): TCellPoint;
    function GetCols(i, j: Integer): Integer;
    function GetRows(i, j: Integer): Integer;
    function GetValues(i, j: Integer): String;
  public
    constructor Create(const AFileName: String; AEncoding: TEncoding);
    destructor Destroy; override;

    property Properties: TGeneralInfo read FProperties;
    property CellCnt: Integer read FCellCnt;
    property Cells[Idx: Integer]: TSection read GetSections; default;
    property BlockType[Idx: Integer]: TBlockType read GetBlockType;
    property Material[Idx: Integer]: TCriteriaMaterial read GetMaterial;
    property FixedID[Idx: Integer]: Integer read GetFixedID;
    property PointCnt[Idx: Integer]: Integer read GetPointCnt;
    property Points[i, j: Integer]: TCellPoint read GetPoints;
    property Cols[i, j: Integer]: Integer read GetCols;
    property Rows[i, j: Integer]: Integer read GetRows;
    property Values[i, j: Integer]: String read GetValues;
  end;

{ TQFmtBuf }
constructor TQFmtBuf.Create(const AFileName: string; AEncoding: TEncoding);
var
  LIdx, i: Integer;
  LBuf: TStringList;
begin
  FSections := TCollections.CreateDictionary<Integer, TSection>([doOwnsValues]);
  FSecNames := TStringList.Create;

  FIni := TMemIniFile.Create(AFileName, AEncoding);
  FCellCnt := 0;

  FIni.ReadSections(FSecNames);

  LBuf := TStringList.Create;
  try
    for i := 0 to FSecNames.Count -1 do
    begin
      LBuf.Clear;
      if FSecNames[i].Equals(SGeneralInfo) then
      begin
        FIni.ReadSectionValues(SGeneralInfo, LBuf);
        FProperties := TGeneralInfo.Create(LBuf);
      end
      else if FSecNames[i].StartsWith(SCell) then
      begin
        LIdx := FSecNames[i].Substring(SCell.Length).ToInteger;
        FIni.ReadSectionValues(SCell + LIdx.ToString, LBuf);
        FSections.Add(LIdx, TSection.Create(LBuf));
        Inc(FCellCnt);
      end;
    end;
  finally
    FreeAndNil(LBuf);
  end;
end;

destructor TQFmtBuf.Destroy;
begin
  FreeAndNil(FSecNames);
  FreeAndNil(FIni);

  inherited;
end;

function TQFmtBuf.GetBlockType(Idx: Integer): TBlockType;
begin
  Result := FSections[Idx].FBlock;
end;

function TQFmtBuf.GetCols(i, j: Integer): Integer;
begin
  Result := Points[i, j].Col;
end;

function TQFmtBuf.GetFixedID(Idx: Integer): Integer;
begin
  Result := FSections[Idx].FFixedID;
end;

function TQFmtBuf.GetMaterial(Idx: Integer): TCriteriaMaterial;
begin
  Result := FSections[Idx].FMaterial;
end;

function TQFmtBuf.GetPointCnt(Idx: Integer): Integer;
begin
  Result := FSections[Idx].PointCnt;
end;

function TQFmtBuf.GetPoints(i, j: Integer): TCellPoint;
begin
  Result := FSections[i].FPoints[j];
end;

function TQFmtBuf.GetRows(i, j: Integer): Integer;
begin
  Result := Points[i, j].Row;
end;

function TQFmtBuf.GetSections(Idx: Integer): TSection;
begin
  Result := FSections[Idx];
end;

function TQFmtBuf.GetValues(i, j: Integer): String;
begin
  Result := Points[i, j].Value;
end;

{ TQFmtBuf.TSection }

constructor TQFmtBuf.TSection.Create(const ASrc: TStringList);
var
  i: Integer;
begin
  FList := TStringList.Create;
  FPoints := TCollections.CreateList<TCellPoint>;

  FList.Assign(ASrc);
  FFixedID := -1;
  FMaterial := cmNone;
  for i := 0 to FList.Count -1 do
    if FList.KeyNames[i].StartsWith('Point') then
      FPoints.Add( TCellPoint.Create(FList.S[ FList.KeyNames[i] ]) )
    else if FList.KeyNames[i].Equals('CriteriaType') then
      FMaterial := TCriteriaMaterial.Create( IfThen( FList.I['CriteriaType'] = 8, 2, FList.I['CriteriaType']))
    else if FList.KeyNames[i].Equals('FixedID') then
      FFixedID := FList.I['FixedID']
    else if FList.KeyNames[i].Equals('ID') then
      FID := FList.S['ID']
    else if FList.KeyNames[i].Equals('BlockType') then
      FBlock := TBlockType.Create( FList.I['BlockType'] );
end;

destructor TQFmtBuf.TSection.Destroy;
begin
  FreeAndNil(FList);

  inherited;
end;

function TQFmtBuf.TSection.GetPointCnt: Integer;
begin
  Result := FPoints.Count
end;

{ TBlockTypeHelper }

class function TBlockTypeHelper.Create(const AValue: Integer): TBlockType;
begin
  Result := TEnum.Parse<TBlockType>(AValue)
end;

function TBlockTypeHelper.ToInteger: Integer;
begin
  Result := Integer(Self);
end;

function TBlockTypeHelper.ToString: String;
begin
  Result := TEnum.GetName<TBlockType>(Self)
end;

{ TCellPoint }

constructor TCellPoint.Create(const ASrc: String);
var
  LSrc: TArray<String>;
begin
  LSrc := ASrc.Split(['|']);
  Assert(Length(LSrc) = 4, 'Invalid parameter: '+ ASrc);

  Col := LSrc[0].ToInteger -1;
  Row := LSrc[1].ToInteger -1;
  FIdx := LSrc[2].ToInteger -1;
  Value := LSrc[3];
end;

constructor TCellPoint.Create(const ACol, ARow, AIdx: Integer; const AValue: String);
begin
  Col := ACol;
  Row := ARow;
  FIdx := AIdx;
  Value := AValue;
end;

function TCellPoint.GetValueAssigned: Boolean;
var
  LValue: Single;
begin
  Result := Single.TryParse(Value, LValue) or Value.ToLower.Equals('over')
end;

function TCellPoint.MatPoint: TMatPoint;
begin
  Result := TMatPoint.Create(Col, Row);
end;

{ TGeneralInfo }

constructor TGeneralInfo.Create(const ARunDate: TDateTime; const ARunNumber, AKitBatchNumber, AOperator: String);
begin
  FRunDate        := ARunDate;
  FRunNumber      := ARunNumber;
  FKitBatchNumber := AKitBatchNumber;
  FOperator       := AOperator;
end;

procedure TGeneralInfo.Assign(const AValue: TGeneralInfo);
begin
  Self := AValue;
  FHasValue := not FRunNumber.IsEmpty or not FKitBatchNumber.IsEmpty or not FOperator.IsEmpty;
end;

procedure TGeneralInfo.Clear;
begin
  Self := System.Default(TGeneralInfo);
  FHasValue := False;
end;

constructor TGeneralInfo.Create(const ASectionValues: TStringList);
var
  LList: TStringList absolute ASectionValues;
  i: Integer;
  function QDateStrToDateTime(const AValue: String; const ASeperator: Char): TDateTime;
  var
    LSrc: TArray<String>;
  begin
    LSrc := AValue.Split(ASeperator);
    if Length(LSrc) = 3 then
      Result := TDateTime.Create(LSrc[2].ToInteger, LSrc[1].ToInteger, LSrc[0].ToInteger)
    else
      Result := MinDateTime;
  end;
begin
  for i := 0 to LList.Count -1 do
    if LList.KeyNames[i].StartsWith('CurveType') then
      FCurveType := LList.I['CurveType']
    else if LList.KeyNames[i].Equals('RunNumber') then
      FRunNumber := LList.S['RunNumber']
    else if LList.KeyNames[i].Equals('KitBatchNumber') then
      FKitBatchNumber := LList.S['KitBatchNumber']
    else if LList.KeyNames[i].Equals('Operator') then
      FOperator := LList.S['Operator']
    else if LList.KeyNames[i].Equals('DataRows') then
      FDataRows := LList.I['DataRows']
    else if LList.KeyNames[i].Equals('DataCols') then
      FDataCols := LList.I['DataCols']
    else if LList.KeyNames[i].Equals('DataRows') then
      FDataRows := LList.I['DataRows']
    else if LList.KeyNames[i].Equals('DecimalSeperator') then
      FDecimalSeparator := LList.C['DecimalSeperator']
    else if LList.KeyNames[i].Equals('DateSeperator') then
      FDateSeparator := LList.C['DateSeperator'];

  FRunDate := QDateStrToDateTime(LList.S['RunDate'], FDateSeparator);
end;

class function TGeneralInfo.Default: TGeneralInfo;
begin
  with Result do
  begin
    FCurveType := 0;
    FDataRows := 9;
    FDataCols := 13;
  end;
end;

function TGeneralInfo.GetAsRunDateStr: String;
var
  y, m, d: Word;
begin
  if not HasValue or (FRunDate = MinDateTime) or FRunDate.IsSameDay(0) then
    Exit('');

  DecodeDate(RunDate, y, m, d);
  Result := Format('%.2d%s%.2d%s%.4d', [d, FormatSettings.DateSeparator, m, FormatSettings.DateSeparator, y]);
end;

function TGeneralInfo.GetDateSeparator: Char;
begin
  Result := FDateSeparator;
  if Result = #0 then
    Result := FormatSettings.DateSeparator;
end;

procedure TGeneralInfo.GetKitBatchNumber(const Value: String);
begin
  if not FHasValue then
    FHasValue := True;

  FKitBatchNumber := Value;
end;

procedure TGeneralInfo.SetInteger(const AIdx, AValue: Integer);
begin
  if not FHasValue then
    FHasValue := True;

  case AIdx of
    NDataRows      : FDataRows  := AValue;
    NDataCols      : FDataCols  := AValue;
  end;
end;

procedure TGeneralInfo.SetOperator(const Value: String);
begin
  if not FHasValue then
    FHasValue := True;
  FOperator := Value;
end;

procedure TGeneralInfo.SetRunDate(const Value: TDateTime);
begin
  if not FHasValue then
    FHasValue := True;

  FRunDate := Value;
end;

procedure TGeneralInfo.SetRunNumber(const Value: String);
begin
  if not FHasValue then
    FHasValue := True;

  FRunNumber := Value;
end;

procedure TGeneralInfo.WriteFmt(const ABuf: TTextWriter; const AContainProperty: Boolean);
const
  CurveTypeFmt = 'CurveType=%d';
  RunDateFmt          = 'RunDate=%s';
  RunNumberFmt        = 'RunNumber=%s';
  KitBatchNumberFmt   = 'KitBatchNumber=%s';
  OperatorFmt         = 'Operator=%s';
  DataRowsFmt         = 'DataRows=%d';
  DataColsFmt         = 'DataCols=%d';
  DecimalSeperatorFmt = 'DecimalSeperator=%s';
  DateSeperatorFmt    = 'DateSeperator=%s';
begin
  ABuf.WriteLine('[GeneralInfo]');
  ABuf.WriteLine(CurveTypeFmt, [CurveType]);
  if HasValue and AContainProperty then
  begin
    ABuf.WriteLine(RunDateFmt         , [AsRunDateStr]);
    ABuf.WriteLine(RunNumberFmt       , [RunNumber]);
    ABuf.WriteLine(KitBatchNumberFmt  , [KitBatchNumber]);
    ABuf.WriteLine(OperatorFmt        , [&Operator]);
    ABuf.WriteLine(DataRowsFmt        , [DataRows]);
    ABuf.WriteLine(DataColsFmt        , [DataCols]);
    ABuf.WriteLine(DecimalSeperatorFmt, [FormatSettings.DecimalSeparator]);
    ABuf.WriteLine(DateSeperatorFmt   , [FormatSettings.DateSeparator]);
  end;
end;

{ TCriteriaTypeHelper }

class function TCriteriaMaterialHelper.Create(const AValue: Integer): TCriteriaMaterial;
begin
  Result := TEnum.Parse<TCriteriaMaterial>(AValue)
end;

function TCriteriaMaterialHelper.GetMaterialString(Index: Integer): String;
begin
  case Self of
    cmNil_Antigen:
    begin
      Assert(InRange(Index, 0, 1), 'Out of index');
      case Index of
        0: Result := Index.ToString + 'N';
        1: Result := Index.ToString + 'A';
      end;
    end;

    cmNil_Antigen_Mitogen:
    begin
      Assert(InRange(Index, 0, 2), 'Out of index');
      case Index of
        0: Result := Index.ToString + 'N';
        1: Result := Index.ToString + 'A';
        2: Result := Index.ToString + 'M';
      end;
    end;

//    cmGold_Nil_ESAT6_Cfp9_Mitogen: ;
//    cmNil_CMV_Mitogen: ;
//    cmNil_CMV_Mitogen_DilCMV_DilMit: ;
//    cmNil_DilutedCMV_DilutedMitogen: ;
//    cmNil_CMV_Mitogen_DilutedCMV: ;
//    cmNil_DilutedCMV_Mitogen: ;
    cmStandard:
    begin
      Assert(InRange(Index, 0, 3), 'Out of index');
      Result := 'S'+ Index.ToString;
    end;
  end;
end;

function TCriteriaMaterialHelper.ToItemCnt: Integer;
begin
  Result := -1;
  case Self of
    cmNil_Antigen: Result := 2;
    cmNil_Antigen_Mitogen: Result := 3;
    cmStandard: Result := 1;
    cmNone: Result := -1;
  else
    Assert(False, 'Handled code does not exists');
  end;
end;

function TCriteriaMaterialHelper.PointLen: Integer;
begin
  Result := -1;
  case Self of
    cmNil_Antigen: Result := 2;
    cmNil_Antigen_Mitogen: Result := 3;
    cmStandard: Result := 4;
    cmNone: Result := 0;
  end;
  Assert(Result > -1, 'Handled code not eixists');
end;

function TCriteriaMaterialHelper.DefaultStdRange(const c, r: Integer): Boolean;
begin
  Result := False;
  case Self of
    cmNil_Antigen: Result := TDefaultMaterial.IsM2StdRange(c, r);
    cmNil_Antigen_Mitogen: Result := TDefaultMaterial.IsM3StdRange(c, r);
  else
    Assert(False, 'Exception occured when process the ' + ToString);
  end;
end;

class function TCriteriaMaterialHelper.CreateByItemCnt(const AValue: Integer): TCriteriaMaterial;
begin
  case AValue of
    2: Result := cmNil_Antigen;
    3: Result := cmNil_Antigen_Mitogen;
    4: Result := cmStandard;
  else
    Result := cmNone;
  end;
end;

function TCriteriaMaterialHelper.DefaultBlock(const c, r: Integer): TBlockType;
begin
  if DefaultStdRange(c, r) then
    Result := btStandard
  else
    Result := btMaterial;
end;

function TCriteriaMaterialHelper.DefaultIdx(const c, r: Integer): Integer;
begin
  Result := -1;
  case Self of
    cmNil_Antigen:
      if DefaultStdRange(c, r) then
        Result := IfThen(c > 5, 1)
      else
        Result := (c * 4) + (r div 2) - IfThen((c >= 5), 2 + IfThen(c > 5, 2));

    cmNil_Antigen_Mitogen:
      if DefaultStdRange(c, r) then
        Result := IfThen(c > 3, 1 + IfThen(c > 4, 1 + IfThen(c > 5, 1)))
      else
        Result := IfThen(c < 3, (c div 3), 4 + 8 * ((c div 3) -1)) + r;
  end;
  Assert(Result > -1, 'Exception occured when process the ' + ToString);
end;

function TCriteriaMaterialHelper.DefaultMaterial(const c, r: Integer): TCriteriaMaterial;
begin
  Result := IfThen(DefaultStdRange(c, r), cmStandard, Self);
end;

function TCriteriaMaterialHelper.ToBlock: TBlockType;
begin
  Result := btEmpty;
  case Self of
    cmNil_Antigen        ,
    cmNil_Antigen_Mitogen: Result := btMaterial;
    cmStandard           : Result := btStandard;
    cmNone               : Result := btEmpty;
  end;
end;

function TCriteriaMaterialHelper.ToInteger: Integer;
begin
  Result := 0;
  case Self of
    cmNil_Antigen: Result := 0;
    cmNil_Antigen_Mitogen: Result := 1;
    cmStandard: Result := 8;
  end;
end;

function TCriteriaMaterialHelper.DefaultPointIdx(const c, r: Integer): Integer;
begin
  Result := -1;
  case Self of
    cmNil_Antigen: Result := IfThen(DefaultStdRange(c, r), r mod 4, r mod 2);
    cmNil_Antigen_Mitogen: Result := IfThen(DefaultStdRange(c, r), r mod 4, c mod 3);
    cmStandard: Result := r mod 4;
  else
    Assert(False, 'Exception occured when process the ' + ToString);
  end;
end;

function TCriteriaMaterialHelper.DefaultPointsLen(const c, r: Integer): Integer;
begin
  Result := -1;
  case Self of
    cmNil_Antigen: Result := IfThen(DefaultStdRange(c, r), 4, 2);
    cmNil_Antigen_Mitogen: Result := IfThen(DefaultStdRange(c, r), 4, 3);
  else
    Assert(False, 'Exception occured when process the ' + ToString);
  end;
end;

function TCriteriaMaterialHelper.ToString: String;
begin
  Result := TEnum.GetName<TCriteriaMaterial>(Self);
end;

{ TCellPointSet }

procedure TSectionCell.AddPoint(AValue: TCellPoint);
begin
  AValue.FIdx := Length(FPoints) +1;
  FPoints := FPoints + [AValue];
end;

constructor TSectionCell.CreateFromSection(const AIdx: Integer; ASection: TStrings; const AFree: Boolean);
begin
  Assert(ASection.KeyExists('Point0'), 'Invalid Paramerter: Point0 is not exists');
  Assert(ASection.KeyExists('BlockType'), 'Invalid Paramerter: BlockType is not exists');

  FPoints := FPoints + [TCellPoint.Create(ASection.S['Point0'])];
  if ASection.KeyExists('Point1') then
    FPoints := FPoints + [TCellPoint.Create(ASection.S['Point1'])];
  if ASection.KeyExists('Point2') then
    FPoints := FPoints + [TCellPoint.Create(ASection.S['Point2'])];
  if ASection.KeyExists('Point3') then
    FPoints := FPoints + [TCellPoint.Create(ASection.S['Point3'])];

  if ASection.KeyExists('CriteriaType') then
    CriteriaMaterial := TCriteriaMaterial.Create(ASection.I['CriteriaType']);

  if ASection.KeyExists('FixedID') then
    FixedID := ASection.I['FixedID'];

  if ASection.KeyExists('ID') then
    ID := ASection.S['ID'];

  BlockType := TBlockType.Create(ASection.I['BlockType']);

  if AFree then
    ASection.Free;
end;

function TSectionCell.AddPoint(const c, r: Integer; AValue: String): TCellPoint;
var
  LIdx: Integer;
begin
  LIdx := IfThen(BlockType = btEmpty, -1, Length(FPoints));
  FPoints := FPoints + [TCellPoint.Create(c, r, LIdx, AValue)];
  Result := FPoints[Length(FPoints) -1]
end;

procedure TSectionCell.ClearPoint;
begin
  FPoints := [];
end;

constructor TSectionCell.Create(const AFixedID: Integer; const ABlock: TBlockType; AMaterial: TCriteriaMaterial; const APoint: TCellPoint);
begin
  Create(AFixedId, ABlock, AMaterial);
  AddPoint(APoint);
end;

constructor TSectionCell.Create(const AFixedID: Integer; const ABlock: TBlockType; AMaterial: TCriteriaMaterial);
begin
  CriteriaMaterial := AMaterial;
  BlockType := ABlock;
  FixedID := IfThen(BlockType <> btStandard, AFixedID);
  ID := IfThen(BlockType = btMaterial, 'ID ') + (AFixedID).ToString;
  ClearPoint;
end;

constructor TSectionCell.CreateEmpty(const AFixedIdx, c, r: Integer; const AValue: String);
begin
  FixedID := AFixedIdx;
  CriteriaMaterial := cmNone;
  BlockType := btEmpty;
  ClearPoint;
  AddPoint(c, r, AValue);
end;

function TSectionCell.GetPointLen: Integer;
begin
  Result := Length(FPoints);
end;

function TSectionCell.PointIdxOf(const c, r: Integer): Integer;
var

  LPoint: TCellPoint;
  i: Integer;
begin
  Result := -1;
  for i := 0 to PointLen -1 do
  begin
    LPoint := FPoints[i];
    if (LPoint.Col = c) and (LPoint.Row = r) then
      Exit(i);
  end;
end;

procedure TSectionCell.SetPointsByMat(MatPoint: TMatPoint; const Value: TCellPoint);
var
  i: Integer;
begin
  i := PointIdxOf(MatPoint.c, MatPoint.r);
  if i > -1 then
    FPoints[i] := Value;
end;

function TSectionCell.ToDoubleArray: TArray<Double>;
var
  i: Integer;
begin
  Result := [];
  for i := 0 to PointLen -1 do
    if Points[i].ValueAssigned then
    begin
      if Points[i].Value.ToLower.Equals('over') then
        Result := Result + [10.]
      else
        Result := Result + [Points[i].Value.ToDouble];
    end;
end;

function TSectionCell.ToStringArray: TStringDynArray;
var
  i: Integer;
begin
  Result := [];
  for i := 0 to PointLen -1 do
    Result := Result + [Points[i].Value];
end;

function TSectionCell.GetIsEmpty: Boolean;
begin
  Result := BlockType = btEmpty;
end;

function TSectionCell.GetLastPoint: TCellPoint;
begin
  Assert(PointLen > 0, 'CellPoint array is empty');
  Result := Points[PointLen -1];
end;

function TSectionCell.GetMaterialText(PointIdx: Integer): String;
begin
  Result := '';
  case CriteriaMaterial of
    cmNil_Antigen:
    begin
      Assert(InRange(PointIdx, 0, 1), 'Out of index');
      case PointIdx of
        0: Result := FixedID.ToString + 'N';
        1: Result := FixedID.ToString + 'A';
      end;
    end;

    cmNil_Antigen_Mitogen:
    begin
      Assert(InRange(PointIdx, 0, 2), 'Out of index');
      case PointIdx of
        0: Result := FixedID.ToString + 'N';
        1: Result := FixedID.ToString + 'A';
        2: Result := FixedID.ToString + 'M';
      end;
    end;

    cmStandard:
    begin
      Assert(InRange(PointIdx, 0, 3), 'Out of index');
      Result := 'S'+ (FixedID + 1).ToString;
    end;
  end;
end;

function TSectionCell.Value(c, r: Integer): String;
begin
  Result := FPoints[ CriteriaMaterial.DefaultPointIdx(c, r) ].Value;
end;

procedure TSectionCell.Write(const AIdx: Integer; const ABuf: TTextWriter);
const
  SSectionFmt = '[Cell%d]';
  SPointFmt = 'Point%d=%s';
  SCriterialTypeFmt = 'CriteriaType=%d';
  SFixedIDFmt = 'FixedID=%d';
  SIDFmt = 'ID=%s';
  SBlockTypeFmt = 'BlockType=%d';
var
  i: Integer;
  LPoint: TCellPoint;
begin
  Assert(Assigned(ABuf), 'ABuf is not assigned!!');

  ABuf.WriteLine(SSectionFmt, [AIdx]);
  i := 0;
  for LPoint in Points do
  begin
    ABuf.WriteLine(SPointFmt, [i,
      Format('%d|%d|%s|%s', [
        LPoint.Col +1,
        LPoint.Row +1,
        IfThen(BlockType = btEmpty, -1, i +1).ToString,
        LPoint.Value])]);
    Inc(i);
  end;

  if not IsEmptyBlock then
  begin
    ABuf.WriteLine(SCriterialTypeFmt, [CriteriaMaterial.ToInteger]);
    ABuf.WriteLine(SFixedIDFmt, [IfThen(BlockType = btMaterial, AIdx +1)]);
    ABuf.WriteLine(SIDFmt, [IfThen(BlockType = btStandard, 'Std ')+ ID]);
  end;
  ABuf.WriteLine(SBlockTypeFmt, [BlockType.ToInteger]);
  if BlockType <> btEmpty then
    ABuf.WriteLine;
end;

{ TDefaultMaterial }

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

{ TRawdataContainer }

procedure TRawdataContainer.AssignDefault(const AMaterial: TCriteriaMaterial);
var
  c, r: Integer;
  LStdRange: Boolean;
  di: Integer;
  LMaterial: TCriteriaMaterial;
  LBlock: TBlockType;
  LList: TSectionCellContainer;
begin
  TMonitor.Enter(FLock);
  try
    BuildRawData;
    FMtrls.Clear;
    FStds.Clear;
    for r := 0 to NRowCnt -1 do
      for c := 0 to FColCount -1 do
      begin
        di := AMaterial.DefaultIdx(c, r);
        LStdRange := AMaterial.DefaultStdRange(c, r);
        LBlock := AMaterial.DefaultBlock(c, r);
        LList := IfThen(LStdRange, FStds, FMtrls);
        if LList.Contains(di) then
          LList.AddPoint(di, FRawData.MatPoints[c, r])
        else
        begin
          LMaterial := AMaterial.DefaultMaterial(c, r);
          LList.AddSectionCell(di, TSectionCell.Create(di +1, LBlock, LMaterial, FRawData.MatPoints[c, r]));
        end;
        FBlockTypes[c, r] := LBlock;
        FCellIdx[c, r] := di;
      end;
    // 기본형식은, AMaterial.DefaultIdx(c, r); 를 통해 정렬된 순서가 c, r에 따라 비순차적으로 들어간다.
    //   때문에 FMtrls.FOrderedByAddedSeq를 Sorting하면, 순서가 보장된다.
    FMtrls.FOrderedByAddedSeq.Sort;
    FRawData.Clear;
  finally
    TMonitor.Exit(FLock);
  end;
  DoChange;
end;

procedure TRawdataContainer.AssignFmtSet(const AValue: TFormatSettings);
begin
  FProperties.DecimalSeparator := AValue.DecimalSeparator;
  FProperties.DateSeparator := AValue.DateSeparator;
end;

procedure TRawdataContainer.AssignManual(c, r: Integer; const AMaterial: TCriteriaMaterial; const ADir: TMaterialDirection);
var
  LPt: TMatPoint;
  di, i: Integer;
  LCell: TSectionCell;
  LList: TSectionCellContainer;
begin
  if FRawData.Count = 0 then
    Exit;

  TMonitor.Enter(FLock);
  try
    CodeSite.EnterMethod(Self, 'AssignManual');
    CodeSite.Send('%d, %d: %s, %s, Loop: %d', [c, r, AMaterial.ToString, ADir.ToString, AMaterial.PointLen]);
    i := 0;
    LList := IfThen(AMaterial.ToBlock = btMaterial, FMtrls, FStds);
    di := LList.Count;
    if LList.Contains(di) then
      Exit;
    LList.AddSectionCell(di, TSectionCell.Create(di +1, AMaterial.ToBlock, AMaterial));
    while i < AMaterial.PointLen do
    begin
      LPt := TMatPoint.Create(c, r);
      CodeSite.Send('MatPoint: %d, %d', [LPt.c, LPt.r]);
      if FRawData.FindEmptyPoint(ADir, LPt, LCell) then
      begin
        LList.AddPoint(di, LCell.LastPoint);
        FBlockTypes[LPt.c, LPt.r] := AMaterial.ToBlock;
        FCellIdx[LPt.c, LPt.r] := di;
        Inc(c, IfThen(ADir = mdHorizontal, 1));
        Inc(r, IfThen(ADir = mdVeritical, 1));
      end;
      Inc(i);
    end;
    CodeSite.ExitMethod(Self, 'AssignManual');
  finally
    TMonitor.Exit(FLock);
  end;
  DoChange;
end;

procedure TRawdataContainer.AssignProperties(const AProperty: TGeneralInfo);
begin
  FProperties := AProperty;
  if Assigned(FOnPropertyChange) then
    FOnPropertyChange;
end;

procedure TRawdataContainer.AssignRandom(const APoint: TMatPoint; const AMaterial: TCriteriaMaterial);
var
  i: Integer;
  LCell: TSectionCell;
  LList: TSectionCellContainer;
begin
  if FRawData.Count = 0 then
    Exit;

  TMonitor.Enter(FLock);
  try
    CodeSite.EnterMethod(Self, 'AssignManualRandom');
    LList := IfThen(AMaterial.ToBlock = btMaterial, FMtrls, FStds);
    i := LList.Count -1;
    if (LList.Count = 0) or (LList.Last.PointLen = AMaterial.PointLen) then
    begin
      i := LList.Count;
      LList.AddSectionCell(i, TSectionCell.Create(i, AMaterial.ToBlock, AMaterial));
    end;

    if FRawData.Extract(APoint, LCell) then
    begin
      CodeSite.Send('[%d] %d, %d: %s', [i, APoint.c, APoint.r, AMaterial.ToString]);
      LList.AddPoint(i, LCell.LastPoint);
      FBlockTypes[APoint.c, APoint.r] := AMaterial.ToBlock;
      FCellIdx[APoint.c, APoint.r] := i;
    end;
    CodeSite.ExitMethod(Self, 'AssignManualRandom');
  finally
    TMonitor.Exit(FLock);
  end;
end;

procedure TRawdataContainer.BuildSeed(const ASrc: TStringList);
var
  LCols: TStringList;
  c, r: Integer;
begin
  LCols := TStringList.Create;
  try
    FColCount := -1;
    FSeed.Clear;
    LCols.StrictDelimiter := True;
    LCols.Delimiter := #9;
    for r := 0 to NRowCnt - 1 do
    begin
      LCols.DelimitedText := ASrc[r];
      while LCols.Count > NColCnt do
        LCols.Delete(LCols.Count - 1);

      if FColCount < LCols.Count then
        FColCount := LCols.Count
      else if FColCount <> LCols.Count then
        raise ELogical.Create('The src columns not match');

      FSeed.Add(LCols.DelimitedText);
    end;

    for r := 0 to FSeed.Count -1 do
    begin
      LCols.DelimitedText := FSeed[r];
      for c := 0 to LCols.Count -1 do
        if LCols[c].IsEmpty then
          LCols[c] := 'N/S';
      FSeed[r] := LCols.DelimitedText;
    end;
  finally
    FreeAndNil(LCols);
  end;
end;

procedure TRawdataContainer.BuildRawData;
var
  i, c, r: Integer;
  LCols: TStringList;
  LCell: TSectionCell;
begin
  LCols := TStringList.Create;
  try
    LCols.StrictDelimiter := True;
    LCols.Delimiter := #9;
    FRawData.Clear;
    i := 0;
    for r := 0 to NRowCnt -1 do
    begin
      LCols.DelimitedText := FSeed[r];
      for c := 0 to LCols.Count -1 do
      begin
        LCell := FRawData.AddSectionCell(i, TSectionCell.CreateEmpty(i, c, r, LCols[c]));
        FBlockTypes[c, r] := btEmpty;
        FCellIdx[c, r] := i;
        Inc(i);
      end;
    end;
  finally
    FreeAndNil(LCols);
  end;
  FProperties.DataRows := NRowCnt +1;
  FProperties.DataCols := FColCount +1;
end;

procedure TRawdataContainer.Clear(const AClearWithProp: Boolean);
var
  c, r: Integer;
begin
  FColCount := NColCnt;
  for c := 0 to NColCnt -1 do
    for r := 0 to NRowCnt -1 do
    begin
      FCellIdx[c, r] := -1;
      FBlockTypes[c, r] := btEmpty;
    end;

  if AClearWithProp then
    FProperties.Clear;
  FSeed.Clear;
  FStds.Clear;
  FMtrls.Clear;
  FRawData.Clear;

  DoChange;
end;

procedure TRawdataContainer.ClearStds;
begin
  ClearSectionCells(FStds);
  DoChange;
end;

procedure TRawdataContainer.ClearSamples;
begin
  ClearSectionCells(FMtrls);
  DoChange;
end;

procedure TRawdataContainer.ClearSectionCells(const ADic: TSectionCellContainer);
var
  c,r, si, di: Integer;
  LSectionCell: TSectionCell;
  LPoint: TCellPoint;
begin
  TMonitor.Enter(FLock);
  try
    for si in ADic.Keys do
      if ADic.Extract(si, LSectionCell) then
        for LPoint in LSectionCell.Points do
        begin
          c := LPoint.Col;
          r := LPoint.Row;
          di := r * ColCount + c;
          FRawData.AddSectionCell(di, TSectionCell.Create(-1, btEmpty, cmNone, LPoint));
          FBlockTypes[c, r] := btEmpty;
          FCellIdx[c, r] := di;
        end;
  finally
    TMonitor.Exit(FLock);
  end;
end;

constructor TRawdataContainer.Create;
 begin
  FLock := TObject.Create;
  FSeed := TRawDataSeed.Create;
  FRawData := TRawDataCellContainer.Create;
  FStds := TSectionCellContainer.Create;
  FMtrls := TSectionCellContainer.Create;

  Clear;
end;

destructor TRawdataContainer.Destroy;
begin
  FreeAndNil(FSeed);
  FreeAndNil(FRawData);
  FreeAndNil(FStds);
  FreeAndNil(FMtrls);
  FreeAndNil(FLock);

  inherited;
end;

procedure TRawdataContainer.DoChange;
begin
  if Assigned(FOnChange) then
    FOnChange;
end;

function TRawdataContainer.GetMatColors(c, r: Integer): TColor;
var
  i: Integer;
begin
  Result := TColors.White;
  if not HasData then
    Exit;

  i := FCellIdx[c, r];
  case FBlockTypes[c, r] of
    btEmpty:
      Result := TColors.White;

    btMaterial:
      case FMtrls[i].CriteriaMaterial of
        cmNil_Antigen        :
          Result := TDefaultMaterial.clM2[Min(i, TDefaultMaterial.NM2MaxLen -1)];

        cmNil_Antigen_Mitogen:
          Result := TDefaultMaterial.clM3[Min(i, TDefaultMaterial.NM3MaxLen -1)];
      end;

    btStandard:
      Result := TDefaultMaterial.clMStd[Min(FStds.MatPoints[c,r].Idx -1, TDefaultMaterial.NMStdLen -1)];

  end;
end;

function TRawdataContainer.GetMatCursors(c, r: Integer): TCursor;
begin
  Result := crDefault;
  if HasData then
    case FBlockTypes[c, r] of
      btEmpty:
        Result := IfThen(MatValues[c, r] <> 'N/S', crHandPoint, crNo);

      btMaterial,
      btStandard:
        Result := crDrag;
    end;
end;

function TRawdataContainer.GetMatFontStyles(c, r: Integer): TFontStyles;
begin
  Result := [];
  if HasData then
  begin
    case FBlockTypes[c, r] of
      btEmpty: Result := [];
      btMaterial: Result := [];
      btStandard: Result := [TFontStyle.fsBold, TFontStyle.fsUnderline];
    end;
  end;
end;

function TRawdataContainer.GetMatIDs(c, r: Integer): String;
begin
  Result := '';
  if HasData then
    case FBlockTypes[c, r] of
      //btEmpty: Result := [];
      btMaterial: Result := FMtrls.MatIDs[c, r];
      //btStandard: Result := FStds.IDs[c, r];
    end;
end;

function TRawdataContainer.GetMatMaterial(c, r: Integer): String;
begin
  if not HasData then
    Exit(' ');

  case FBlockTypes[c, r] of
    btEmpty: Result := ' ';
    btMaterial: Result := FMtrls.MatMaterialText[c, r];
    btStandard: Result := FStds.MatMaterialText[c, r];
  end;
end;

function TRawdataContainer.GetMatValues(c, r: Integer): String;
begin
  Result := ' ';
  if HasData then
  begin
    case FBlockTypes[c, r] of
      btEmpty: Result := FRawData.MatPoints[c, r].Value;
      btMaterial: Result := FMtrls.MatPoints[c,r].Value;
      btStandard: Result := FStds.MatPoints[c,r].Value;
    end;
  end;
end;

function TRawdataContainer.GetPointsByIds(Idx: Integer): TArray<TCellPoint>;
begin
  Result := FMtrls.SectionCells[Idx].Points;
end;

function TRawdataContainer.GetStdCount: Integer;
begin
  Result := FStds.Count;
end;

function TRawdataContainer.IDArray: TArray<String>;
var
  k, i: Integer;
begin
  i := 0;
  SetLength(Result, FMtrls.Count);
  for k in FMtrls.Keys do
  begin
    Result[i] := FMtrls.SectionCells[k].ID;
    Inc(i);
  end;
end;

function TRawdataContainer.GetCanCalc: Boolean;
begin
  Result := (FMtrls.Count > 0) and (FStds.Count > 1);
end;

function TRawdataContainer.GetCellCount: Integer;
begin
  Result := FMtrls.Count;
end;

function TRawdataContainer.GetHasData: Boolean;
begin
  Result := FSeed.Count > 0;
end;

function TRawdataContainer.GetHasStd: Boolean;
begin
  Result := FStds.Count > 0;
end;

function TRawdataContainer.GetIDs(Idx: Integer): String;
begin
  Result := FMtrls.SectionCells[Idx].ID;
end;

function TRawdataContainer.MatIDArray: TArray<TMatStrPair>;
begin
  Result := FMtrls.FIDs.ToArray;
end;

procedure TRawdataContainer.LoadFromFile(const AFileName: String);
var
  LBuf: TQFmtBuf;
  LCols: TStringList;
  di, i, j, c, r: Integer;
begin
  TMonitor.Enter(FLock);
  try
    Clear;
    LBuf := TQFmtBuf.Create(AFileName, TEncoding.Unicode);
    try
      FProperties.Assign(LBuf.Properties);
      for i := 0 to LBuf.CellCnt -1 do
      begin
        case LBuf.BlockType[i] of
          btMaterial:
          begin
            di := FMtrls.Count;
            if LBuf.PointCnt[i] = 0 then
              FMtrls.AddSectionCell(di, TSectionCell.Create(LBuf.FixedID[i], LBuf.BlockType[i], LBuf.Material[i]))
            else
            begin
              FMtrls.AddSectionCell(di, TSectionCell.Create(LBuf.FixedID[i], LBuf.BlockType[i], LBuf.Material[i], LBuf.Points[i, 0]));
              FBlockTypes[LBuf.Cols[i, 0], LBuf.Rows[i, 0]] := btMaterial;
              FCellIdx[LBuf.Cols[i, 0], LBuf.Rows[i, 0]] := di;
              for j := 1 to LBuf.PointCnt[i] -1 do
              begin
                FMtrls.AddPoint(di, LBuf.Points[i, j]);
                FBlockTypes[LBuf.Cols[i, j], LBuf.Rows[i, j]] := btMaterial;
                FCellIdx[LBuf.Cols[i, j], LBuf.Rows[i, j]] := di;
              end;
            end;
          end;
          btStandard:
          begin
            di := FStds.Count;
            if LBuf.PointCnt[i] = 0 then
              FStds.AddSectionCell(di, TSectionCell.Create(LBuf.FixedID[i], LBuf.BlockType[i], LBuf.Material[i]))
            else
            begin
              FStds.AddSectionCell(di, TSectionCell.Create(LBuf.FixedID[i], LBuf.BlockType[i], LBuf.Material[i], LBuf.Points[i, 0]));
              FBlockTypes[LBuf.Cols[i, 0], LBuf.Rows[i, 0]] := btStandard;
              FCellIdx[LBuf.Cols[i, 0], LBuf.Rows[i, 0]] := di;
              for j := 1 to LBuf.PointCnt[i] -1 do
              begin
                FStds.AddPoint(di, LBuf.Points[i, j]);
                FBlockTypes[LBuf.Cols[i, j], LBuf.Rows[i, j]] := btStandard;
                FCellIdx[LBuf.Cols[i, j], LBuf.Rows[i, j]] := di;
              end;
            end;
          end;

          btEmpty:
          begin
            di := FRawdata.Count;
            j := 0;
            c := LBuf.Cols[i, j]; Assert(c < 12, 'ccccccccccccc');
            r := LBuf.Rows[i, j]; Assert(r <  8, 'rrrrrrrrr');
            FRawdata.AddSectionCell(di, TSectionCell.CreateEmpty(di, c, r, LBuf.Values[i, j]));
            FBlockTypes[c, r] := btEmpty;
            FCellIdx[c, r] := di;
          end
        end;
      end;
    finally
      FreeAndNil(LBuf);
    end;
  finally
    TMonitor.Exit(FLock);
  end;

  FSeed.Clear;
  LCols := TStringList.Create;
  try
    LCols.StrictDelimiter := True;
    LCols.Delimiter := #9;
    for r := 0 to NRowCnt -1 do
    begin
      LCols.Clear;
      for c := 0 to NColCnt -1 do
        if FCellIdx[c, r] > -1 then
          case FBlockTypes[c, r] of
            btMaterial: LCols.Add(FMtrls.MatPoints[c, r].Value);
            btStandard: LCols.Add(FStds.MatPoints[c, r].Value);
            btEmpty: LCols.Add(FRawdata.MatPoints[c, r].Value);
          end;
      FColCount := LCols.Count;
      FSeed.Add(LCols.DelimitedText);
    end;
    CodeSite.Send(FSeed.Text);
  finally
    LCols.Free;
  end;
  DoChange;
end;

function TRawdataContainer.MtrlArray: TArray<TStringDynArray>;
var
  k: Integer;
begin
  Result := [];
  for k in FMtrls.Keys do
    Result := Result + [FMtrls[k].ToStringArray]
end;

procedure TRawdataContainer.Paste(const ASrc: TStringList; const AColCount: Integer);
begin
  TMonitor.Enter(FLock);
  try
    BuildSeed(ASrc);
    BuildRawData;
  finally
    TMonitor.Exit(FLock);
  end;
  DoChange;
end;

procedure TRawdataContainer.RemoveFmt(c, r: Integer);
var
  di, si: Integer;
  LBlockType: TBlockType;
  LSectionCell: TSectionCell;
  LPoint: TCellPoint;
  LList: TSectionCellContainer;
begin
  CodeSite.EnterMethod(Self, 'RemoveFmt');
  TMonitor.Enter(FLock);
  try
    LBlockType := FBlockTypes[c, r];
    CodeSite.Send('%d, %d', [c, r]);
    case LBlockType of
      btEmpty:;

      btMaterial,
      btStandard:
      begin
        LList := IfThen(LBlockType = btMaterial, FMtrls, FStds);
        si := LList.FIdxs[TMatPoint.Create(c, r)];
        if LList.Extract(si, LSectionCell) then
          for LPoint in LSectionCell.Points do
          begin
            c := LPoint.Col;
            r := LPoint.Row;
            di := r * ColCount + c;
            FRawData.AddSectionCell(di, TSectionCell.Create(-1, btEmpty, cmNone, LPoint));
            FBlockTypes[c, r] := btEmpty;
            FCellIdx[c, r] := di;
          end;
      end;
    end;
  finally
    TMonitor.Exit(FLock);
    CodeSite.ExitMethod(Self, 'RemoveFmt');
  end;
  DoChange;
end;

procedure TRawdataContainer.SaveToFile(const AFileName: String; const AContainsProperty: Boolean);
var
  k, LCellIdx: Integer;
  LBuf: TStreamWriter;
begin
  LBuf := TStreamWriter.Create(AFileName, False, TEncoding.Unicode);
  try
    FProperties.WriteFmt(LBuf, AContainsProperty);

    LCellIdx := 0;
    for k in FMtrls.Keys do
    begin
      FMtrls.SectionCells[k].Write(LCellIdx, LBuf);
      Inc(LCellIdx);
    end;

    for k in FStds.Keys do
    begin
      FStds.SectionCells[k].Write(LCellIdx, LBuf);
      Inc(LCellIdx);
    end;

    for k in FRawData.SortedByPointKeys do
    begin
      FRawData.SectionCells[k].Write(LCellIdx, LBuf);
      Inc(LCellIdx);
    end;
  finally
    FreeAndNil(LBuf);
  end;
end;

procedure TRawdataContainer.SetMatIDs(c, r: Integer; const Value: String);
begin
  if HasData then
    case FBlockTypes[c, r] of
      //btEmpty: Result := [];
      btMaterial: FMtrls.MatIDs[c, r] := Value;
      //btStandard: Result := FStds.IDs[c, r];
    end;
end;

procedure TRawdataContainer.SetMatValues(c, r: Integer; const Value: String);
begin
  if HasData then
  begin
    FSeed.Mats[c, r] := Value;
    case FBlockTypes[c, r] of
      btEmpty: FRawData.MatValues[c, r] := Value;
      btMaterial: FMtrls.MatValues[c, r] := Value;
      btStandard: FStds.MatValues[c, r] := Value;
    end;
    if Assigned(FOnCellChange) then
      FOnCellChange(c, r);
  end;
end;

function TRawdataContainer.StdArray: TArray<TArray<Double>>;
var
  k: Integer;
begin
  Result := [];
  for k in FStds.Keys do
    Result := Result + [FStds[k].ToDoubleArray]
end;

{ TRawdataFormater }

procedure TRawdataFormater.SetEnabled(const Value: Boolean);
begin
  FEnabled := Value;
end;

procedure TRawdataFormater.StepIt;
begin
  Dec(FRemainCnt, FCapacity);
  DoChange;

  if FRemainCnt = 0 then
    Initialize;
end;

procedure TRawdataFormater.DoChange;
begin
  if Assigned(FOnChange) then
    FOnChange(FRemainCnt);
end;

constructor TRawdataFormater.Create;
begin
  Initialize(cmStandard, mdVeritical);
  FEnabled := False;
end;

function TRawdataFormater.GetCanInit: Boolean;
begin
  Result := False;
  case FSeed of
    cmNil_Antigen: Result := FRemainCnt in [0, 2];
    cmNil_Antigen_Mitogen: Result := FRemainCnt in [0, 3];
    cmStandard: Result := FRemainCnt in [0, 4];
  end;
end;

procedure TRawdataFormater.Initialize;
begin
  case FSeed of
    cmNil_Antigen: FRemainCnt := 2;
    cmNil_Antigen_Mitogen: FRemainCnt := 3;
    cmStandard: FRemainCnt := 4;
  else
    Assert(False, 'Handled code not exists!!');
  end;
  FCapacity := IfThen(FDir = mdRandom, 1, FRemainCnt);
  FEnabled := True;
  DoChange;
end;

procedure TRawdataFormater.Initialize(const ASeed: TCriteriaMaterial; const ADir: TMaterialDirection);
begin
  CodeSite.EnterMethod(Self, 'Initialize');
  CodeSite.Send('Material', ASeed.ToString);
  CodeSite.Send('Direction', ADir.ToString);
  FSeed := ASeed;
  FDir := ADir;
  Initialize;
  CodeSite.ExitMethod(Self, 'Initialize');
end;

{ TMaterialDirectionHelper }

class function TMaterialDirectionHelper.Create(const AValue: Integer): TMaterialDirection;
begin
  Assert(InRange(AValue, Low(TMaterialDirection).ToInteger, High(TMaterialDirection).ToInteger), 'Invalid Parameter!!');
  Result := TEnum.Parse<TMaterialDirection>(AValue);
end;

function TMaterialDirectionHelper.ToInteger: Integer;
begin
  Result := Integer(Self);
end;

function TMaterialDirectionHelper.ToString: String;
begin
  Result := TEnum.GetName<TMaterialDirection>(Self)
end;

{ TSectionCellList }

function TSectionCellContainer.AddPoint(const ADstIdx: Integer; const AValue: TCellPoint): TSectionCell;
var
  LCell: TSectionCell;
  LKey: TMatPoint;
  LPoint: TCellPoint;
begin
  LCell := FCells[ADstIdx];
  LCell.AddPoint(AValue);
  FCells[ADstIdx] := LCell;
  Result := LCell;

  LPoint := Result.LastPoint;
  LKey := LPoint.MatPoint ;
  FPoints.AddOrSetValue(LKey, LPoint);
  FIdxs.AddOrSetValue(LKey, ADstIdx);
  FMtrlText.AddOrSetValue(Result.BlockType, Result.FixedID, Result.PointLen -1, LPoint);
  if LCell.PointLen = 1 then
  begin
    FIDs.Add(LKey, Result.ID);
    FOrderedByAddedSeq.Add(ADstIdx);
  end;
//  if not FIDs.ContainsKey(LKey) then
//    FIDs.Add(LKey, Result.ID);
end;

function TSectionCellContainer.AddSectionCell(const ADstIdx: Integer; const AValue: TSectionCell): TSectionCell;
var
  LKey: TMatPoint;
  LPoint: TCellPoint;
  LPointLen: Integer;
begin
  FCells.Add(ADstIdx, AValue);
  Result := FCells[ADstIdx];
  LPointLen := Result.PointLen;
  if LPointLen > 0 then
  begin
    LPoint := Result.LastPoint;
    LKey := LPoint.MatPoint;
    FPoints.Add(LKey, LPoint);
    FIdxs.Add(LKey, ADstIdx);
    FMtrlText.Add(AValue.BlockType, AValue.FixedID, LPointLen -1, LPoint);
    FIDs.Add(LKey, Result.ID);
    FOrderedByAddedSeq.Add(ADstIdx);
  end;
end;

procedure TSectionCellContainer.Clear;
begin
  FCells.Clear;
  FPoints.Clear;
  FMtrlText.Clear;
  FIdxs.Clear;
  FIDs.Clear;
  FOrderedByAddedSeq.Clear;
end;

constructor TSectionCellContainer.Create;
begin
  FCells := TCollections.CreateDictionary<Integer, TSectionCell>;
  FPoints := TCollections.CreateDictionary<TMatPoint, TCellPoint>;
  FIdxs := TCollections.CreateDictionary<TMatPoint, Integer>;
  FIds :=  TCollections.CreateDictionary<TMatPoint, String>;
  FMtrlText := TMaterialTextDic.Create;
  FOrderedByAddedSeq := TCollections.CreateList<Integer>;
end;

destructor TSectionCellContainer.Destroy;
begin
  FreeAndNil(FMtrlText);

  inherited;
end;

function TSectionCellContainer.Contains(const AIdx: Integer): Boolean;
begin
  Result := FCells.ContainsKey(AIdx)
end;

function TSectionCellContainer.Extract(const AIdx: Integer; var AValue: TSectionCell): Boolean;
var
  LPoint: TCellPoint;
begin
  Result := FCells.ContainsKey(AIdx);
  if Result then
  begin
    AValue := FCells.ExtractPair(AIdx).Value;
    for LPoint in AValue.Points do
    begin
      FIdxs.Remove(LPoint.MatPoint);
      FPoints.Remove(LPoint.MatPoint);
      FMtrlText.Remove(LPoint.MatPoint);
      if FIDs.ContainsKey(LPoint.MatPoint) then
        FIDs.Remove(LPoint.MatPoint);
      FOrderedByAddedSeq.Extract(AIdx);
    end;
  end;
end;

function TSectionCellContainer.GetCount: Integer;
begin
  Result := FCells.Count;
end;

function TSectionCellContainer.GetMatIDs(c, r: Integer): String;
var
  LMatPoint: TMatPoint;
begin
  Result := ' ';
  LMatPoint := TMatPoint.Create(c, r);
  if FIDs.ContainsKey(LMatPoint) then
    Result := FIDs[LMatPoint];
end;

function TSectionCellContainer.GetSectionCells(Idx: Integer): TSectionCell;
begin
  Result := FCells[Idx]
end;

function TSectionCellContainer.GetLast: TSectionCell;
begin
  Result := SectionCells[Count -1];
end;

function TSectionCellContainer.GetMatMaterialText(c, r: Integer): String;
begin
  Result := FMtrlText[c, r]
end;

function TSectionCellContainer.GetMatPoints(c, r: Integer): TCellPoint;
begin
  Result := FPoints[TMatPoint.Create(c, r)];
end;

function TSectionCellContainer.GetMatValues(c, r: Integer): String;
begin
  Result := FPoints[TMatPoint.Create(c, r)].Value
end;

function TSectionCellContainer.Keys: TArray<Integer>;
begin
  //Result := FCells.Keys.ToArray;
  //FOrderedByAddedSeq.Sort;
  Result := FOrderedByAddedSeq.ToArray;
end;

procedure TSectionCellContainer.SetMatIDs(c, r: Integer; const Value: String);
var
  LMat: TMatPoint;
  LCell: TSectionCell;
  i: Integer;
begin
  LMat := TMatPoint.Create(c, r);
  i := FIdxs[LMat];
  LCell := FCells[i];
  LCell.ID := Value;
  FCells[i] := LCell;
  FIDs[LMat] := Value;
end;

procedure TSectionCellContainer.SetSectionCells(Idx: Integer; const Value: TSectionCell);
begin
  FCells[Idx] := Value;
end;

function TSectionCellContainer.SortedByPointKeys: TArray<Integer>;
var
  LList: ILIst<Integer>;
begin
  LList := TCollections.CreateList<Integer>(TComparer<Integer>.Construct(
    function(const ALeft, ARight: Integer): Integer
    begin
      Result := CompareValue(FCells.Items[ALeft].LastPoint.Col, FCells.Items[ARight].LastPoint.Col);
      if Result = EqualsValue then
        Result := CompareValue(FCells.Items[ALeft].LastPoint.Row, FCells.Items[ARight].LastPoint.Row);
    end));
  LList.AddRange(FCells.Keys.ToArray);
  LList.Sort;
  Result := LList.ToArray;
end;

procedure TSectionCellContainer.SetMatPoints(c, r: Integer; const Value: TCellPoint);
begin
  FPoints[TMatPoint.Create(c, r)] := Value;
end;

procedure TSectionCellContainer.SetMatValues(c, r: Integer; const Value: String);
var
  LMatPoint: TMatPoint;
  LPoint: TCellPoint;
begin
  LMatPoint := TMatPoint.Create(c, r);
  LPoint := FPoints[LMatPoint];
  LPoint.Value := Value;
  FPoints[LMatPoint] := LPoint;
  FCells[FIdxs[LMatPoint]].SetPointsByMat(LMatPoint, LPoint);
end;

{ TMat }

constructor TMatPoint.Create(const ACol, ARow: Integer);
begin
  c := ACol;
  r := ARow;
end;

{ TRawDataCellContainer }

function TRawDataCellContainer.Extract(const AKey: TMatPoint; var AValue: TSectionCell): Boolean;
var
  LIdx: Integer;
begin
  Result := FIdxs.ContainsKey(AKey) and MatPoints[AKey.c, AKey.r].ValueAssigned;
  if Result then
  begin
    LIdx := FIdxs.ExtractPair(AKey).Value;
    AValue := FCells.ExtractPair(LIdx).Value;
    FPoints.Remove(AValue.LastPoint.MatPoint);
    FMtrlText.Remove(AValue.LastPoint.MatPoint);
    FIDs.Remove(AValue.LastPoint.MatPoint);
  end;
end;

function TRawDataCellContainer.FindEmptyPoint(const ADir: TMaterialDirection; var APt: TMatPoint;
  var AValue: TSectionCell): Boolean;
var
  ci, ri: Integer;
begin

  Result := False;
  //CodeSite.EnterMethod(Self, 'FindEmptyPoint');
  try
    case ADir of
      mdVeritical:
      begin
        for ci in [APt.c..11] do
          for ri in [IfThen(APt.c = ci, APt.r)..7] do
          begin
            Result := Extract(TMatPoint.Create(ci, ri), AValue);
            if Result then
            begin
              APt.c := ci;
              APt.r := ri;
              Exit;
            end;
          end;
        for ci in [0..APt.c] do
          for ri in [0..APt.r] do
          begin
            Result := Extract(TMatPoint.Create(ci, ri), AValue);
            if Result then
            begin
              APt.c := ci;
              APt.r := ri;
              Exit;
            end;
          end;
      end;

      mdHorizontal:
      begin
        for ri in [APt.r..7] do
          for ci in [IfThen(APt.r = ri, APt.c)..11] do
          begin
//            CodeSite.Send('iteration: %d, %d', [ci, ri]);
            Result := Extract(TMatPoint.Create(ci, ri), AValue);
            if Result then
            begin
              APt.c := ci;
              APt.r := ri;
              Exit;
            end;
          end;
        for ri in [0..APt.r] do
          for ci in [0..APt.c-1] do
          begin
//            CodeSite.Send('iteration: %d, %d', [ci, ri]);
            Result := Extract(TMatPoint.Create(ci, ri), AValue);
            if Result then
            begin
              APt.c := ci;
              APt.r := ri;
              Exit;
            end;
          end;
      end;

      mdRandom:
        Result := Extract(TMatPoint.Create(APt.c, APt.r), AValue);
    end;
  finally
//    CodeSite.Send('MatPoint: %d, %d', [APt.c, APt.r]);
//    CodeSite.ExitMethod(Self, 'FindEmptyPoint', Result);
  end;
end;

{ TMaterialTextDic }

procedure TMaterialTextDic.Add(const ABlock: TBlockType; const AFixedId, APointIdx: Integer; const APoint: TCellPoint);
const
  SMtrlText: array[0..2] of String = ('N', 'A', 'M');
var
  LKey: TMatPoint;
begin
  LKey := APoint.MatPoint;
  case ABlock of
    btMaterial: FDic.Add(LKey, AFixedId.ToString + SMtrlText[APointIdx]);
    btStandard: FDic.Add(LKey,'S' + (APointIdx + 1).ToString);
    btEmpty: FDic.Add(LKey, ' ');
  end;
end;

procedure TMaterialTextDic.AddOrSetValue(const ABlock: TBlockType; const AFixedId, APointIdx: Integer; const APoint: TCellPoint);
const
  SMtrlText: array[0..2] of String = ('N', 'A', 'M');
var
  LKey: TMatPoint;
begin
  LKey := APoint.MatPoint;
  case ABlock of
    btMaterial: FDic.AddOrSetValue(LKey, AFixedId.ToString + SMtrlText[APointIdx]);
    btStandard: FDic.AddOrSetValue(LKey,'S' + (APointIdx + 1).ToString);
    btEmpty: FDic.AddOrSetValue(LKey, ' ');
  end;
end;

procedure TMaterialTextDic.Clear;
begin
  FDic.Clear;
end;

constructor TMaterialTextDic.Create;
begin
  FDic := TCollections.CreateDictionary<TMatPoint, String>;
end;

function TMaterialTextDic.GetItems(c, r: Integer): String;
begin
  Result := FDic[TMatPoint.Create(c, r)];
end;

procedure TMaterialTextDic.Remove(const AKey: TMatPoint);
begin
  FDic.Remove(AKey)
end;

{ TRawDataSeed }

function TRawDataSeed.Cols(const r: Integer; var ACols: TStringList): Boolean;
begin
  Result := r < Count;
  if Result then
  begin
    ACols := TStringList.Create;
    ACols.StrictDelimiter := True;
    ACols.Delimiter := #9;
    ACols.DelimitedText := Strings[r];
  end;
end;

function TRawDataSeed.GetMats(c, r: Integer): String;
var
  LCols: TStringList;
begin
  if Cols(r, LCols) then
    try
      Result := IfThen(LCols.Count > c, LCols[c]);
    finally
      FreeAndNil(LCols);
    end;
end;

procedure TRawDataSeed.SetMats(c, r: Integer; const Value: String);
var
  LCols: TStringList;
begin
  if Cols(r, LCols) then
    try
      LCols[c] := Value;
      Strings[r] := LCols.DelimitedText;
    finally
      FreeAndNil(LCols);
    end;
end;

{ TMaterialHelper }

class function TMaterialHelper.Create(const AValue: Integer): TMaterial;
begin
  Assert(AValue in [0..2], 'Invalid parameter. It can have only 0 to 2.');
  Result := TEnum.Parse<TMaterial>(AValue);
end;

class function TMaterialHelper.ToArray: TArray<TMaterial>;
var
  LItem: TMaterial;
begin
  Result := [];
  for LItem := Low(TMaterial) to High(TMaterial) do
    Result := Result + [LItem];
end;

end.
