unit m.block;

interface

uses
  System.Sysutils, System.classes, Spring, System.UITypes, Spring.Collections, System.SyncObjs,
  System.Generics.Collections
  ;

const
  NColCnt = 12;
  NMinColCnt = 2;
  NRowCnt =  8;

type
  TMatArray<T> = array[0..NColCnt -1, 0..NRowCnt -1] of T;

  TMatrix = record
    c,
    r: Integer;
    constructor Create(const ACol, ARow: Integer);
  end;

  TMatDictionary<T> = class(TDictionary<TMatrix, T>)
  end;

type
  TValueKind = (
    vkValue,
    vkOver,
    vkNotSpecified
  );

const
  SValueNotSpecified = 'N/S';
  SValueOver = 'over';

type
  TPoint = record
  private
    FCol, FRow: Integer;
    FIdx: Integer;
    FValue: String;
    function GetAsDoubleValue: Double;
    function GetHasValue: Boolean;
    function GetIsValueOver: Boolean;
    function GetValue: String;
    function GetMatrix: TMatrix;
  public
    constructor CreateFromSectionValue(const ASrc: string);
    constructor CreateFromClipbrd(const ACol, ARow: Integer; const AValue: string);

    function ToString: string;

    property Col: Integer read FCol;
    property Row: Integer read FRow;
    property Idx: Integer read FIdx;
    property Value: String read GetValue;
    property AsDoubleValue: Double read GetAsDoubleValue;
    property HasValue: Boolean read GetHasValue;
    property IsValueOver: Boolean read GetIsValueOver;
    property Mat: TMatrix read GetMatrix;
  end;

  TBlockType = (
    btSample = 0,
    btStandard,
    btNotSpecified
  );
  TBlockTypeHelper = record helper for TBlockType
    function ToInteger: Integer;
    class function Create(const AValue: Integer): TBlockType; static;
  end;

  TCriteriaType = (
    ctNil_Antigen,
    ctNil_Antigen_Mitogen,
    ctStandard
  );
  TCriteriaTypeHelper = record helper for TCriteriaType
    class function Create(const AValue: Integer): TCriteriaType; static;

    function ToInteger: Integer;
    function DefaultIdx(const c, r: Integer): Integer;
    function DefaultStdRange(const c, r: Integer): Boolean;
    function DefaultBlock(const c, r: Integer): TBlockType;
  end;

  TPointIdx = 0..3;
  TBlock = record
  private
    FPoints: TArray<TPoint>;
    FCriteriaType: Nullable<TCriteriaType>;
    FFixedID: Integer;
    FId: string;
    FBlockType: TBlockType;
    FRandomAssigned: Boolean;
    function AddPoint(const AValue: TPoint): Integer; overload;
    constructor CreateFromPoint(const APoint: TPoint);
  private
    function GetPointCnt: Integer;
    function GetPoints(Idx: TPointIdx): TPoint;
    function GetCriteriaType: TCriteriaType;
    function GetFontStyles: TFontStyles;
    function GetCursor: TCursor;
    function GetBgColors(Idx: TPointIdx): TColor;
    function GetMtrlTexts(Index: TPointIdx): String;
    function GetValues(Idx: TPointIdx): String;
    procedure SetId(const Value: String);
    function GetLastPoint: TPoint;
  public const
    SSecCellFmt = 'Cell%d';
    SIdentPoints: array[0..3] of String = ('Point0', 'Point1', 'Point2', 'Point3');
    SIdentCriteriaType = 'CriteriaType';
    SIdentFixedId = 'FixedID';
    SIdentId = 'ID';
    SIdentBlockType = 'BlockType';
  public
    constructor CreateFromClipbrd(const ACol, ARow: Integer; const AValue: string);
    constructor CreateFromSection(const CellIdx: Integer; ASec: TStrings; const AFreeSection: Boolean);

    function AddPointFromBlock(const ASeq: Integer; const ABlocks: TArray<TBlock>): Integer; overload;
    function ToNoSpecifiedBlocks: TArray<TBlock>;

    procedure WriteBuf(const ASeq: Integer; const ABuf: TTextWriter);

    property PointCnt: Integer read GetPointCnt;
    property Points[Idx: TPointIdx]: TPoint read GetPoints; default;
    property LastPoint: TPoint read GetLastPoint;
    property CriteriaType: TCriteriaType read GetCriteriaType;
    property FixedId: Integer read FFixedID;
    property Id: String read FId write SetId;
    property BlockType: TBlockType read FBlockType;
    property MtrlTexts[Idx: TPointIdx]: String read GetMtrlTexts;
    property Values[Idx: TPointIdx]: String read GetValues;
    property BgColors[Idx: TPointIdx]: TColor read GetBgColors;
    property FontStyle: TFontStyles read GetFontStyles;
    property Cursor: TCursor read GetCursor;
  end;

  TSample = class
  public const
    NMStdLen = 4;
    clMStd: array[0..NMStdLen -1] of TColor = (
      $008C8C8C, $00AAAAAA, $00C8C8C8, $00E6E6E6
    );
    matStdM2: array[0..1, 0..3] of TMatrix = (
      ((c: 5; r:0), (c: 5; r:1), (c: 5; r:2), (c: 5; r:3)),
      ((c: 6; r:0), (c: 6; r:1), (c: 6; r:2), (c: 6; r:3))
    );
    matStdM3: array[0..2, 0..3] of TMatrix = (
      ((c: 3; r:0), (c: 3; r:1), (c: 3; r:2), (c: 3; r:3)),
      ((c: 4; r:0), (c: 4; r:1), (c: 4; r:2), (c: 4; r:3)),
      ((c: 5; r:0), (c: 5; r:1), (c: 5; r:2), (c: 5; r:3))
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
    matM2: array[0..NM2Len-1, 0..1] of TMatrix = (
      {$REGION 'matM2 matrix'}
      ((c: 0; r:0), (c: 0; r:1)), ((c: 0; r:2), (c: 0; r:3)), ((c: 0; r:4), (c: 0; r:5)), ((c: 0; r:6), (c: 0; r:7)),
      ((c: 1; r:0), (c: 1; r:1)), ((c: 1; r:2), (c: 1; r:3)), ((c: 1; r:4), (c: 1; r:5)), ((c: 1; r:6), (c: 1; r:7)),
      ((c: 2; r:0), (c: 2; r:1)), ((c: 2; r:2), (c: 2; r:3)), ((c: 2; r:4), (c: 2; r:5)), ((c: 2; r:6), (c: 2; r:7)),
      ((c: 3; r:0), (c: 3; r:1)), ((c: 3; r:2), (c: 3; r:3)), ((c: 3; r:4), (c: 3; r:5)), ((c: 3; r:6), (c: 3; r:7)),
      ((c: 4; r:0), (c: 4; r:1)), ((c: 4; r:2), (c: 4; r:3)), ((c: 4; r:4), (c: 4; r:5)), ((c: 4; r:6), (c: 4; r:7)),
     {((c: 5; r:0), (c: 5; r:1)), ((c: 5; r:2), (c: 5; r:3)),}((c: 5; r:4), (c: 5; r:5)), ((c: 5; r:6), (c: 5; r:7)),
     {((c: 6; r:0), (c: 6; r:1)), ((c: 6; r:2), (c: 6; r:3)),}((c: 6; r:4), (c: 6; r:5)), ((c: 6; r:6), (c: 6; r:7)),
      ((c: 7; r:0), (c: 7; r:1)), ((c: 7; r:2), (c: 7; r:3)), ((c: 7; r:4), (c: 7; r:5)), ((c: 7; r:6), (c: 7; r:7)),
      ((c: 8; r:0), (c: 8; r:1)), ((c: 8; r:2), (c: 8; r:3)), ((c: 8; r:4), (c: 8; r:5)), ((c: 8; r:6), (c: 8; r:7)),
      ((c: 9; r:0), (c: 9; r:1)), ((c: 9; r:2), (c: 9; r:3)), ((c: 9; r:4), (c: 9; r:5)), ((c: 9; r:6), (c: 9; r:7)),
      ((c:10; r:0), (c:10; r:1)), ((c:10; r:2), (c:10; r:3)), ((c:10; r:4), (c:10; r:5)), ((c:10; r:6), (c:10; r:7)),
      ((c:11; r:0), (c:11; r:1)), ((c:11; r:2), (c:11; r:3)), ((c:11; r:4), (c:11; r:5)), ((c:11; r:6), (c:11; r:7))
      {$ENDREGION}
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
    matM3: array[0..NM3Len -1, 0..2] of TMatrix = (
     {$REGION 'matM3 matrix'}
      ((c:0; r:0), (c:1; r:0), (c:2; r:0)),
      ((c:0; r:1), (c:1; r:1), (c:2; r:1)),
      ((c:0; r:2), (c:1; r:2), (c:2; r:2)),
      ((c:0; r:3), (c:1; r:3), (c:2; r:3)),
      ((c:0; r:4), (c:1; r:4), (c:2; r:4)),
      ((c:0; r:5), (c:1; r:5), (c:2; r:5)),
      ((c:0; r:6), (c:1; r:6), (c:2; r:6)),
      ((c:0; r:7), (c:1; r:7), (c:2; r:7)),
                                            {((c:3; r:0), (c:4; r:0), (c:5; r:0)),}
                                            {((c:3; r:1), (c:4; r:1), (c:5; r:1)),}
                                            {((c:3; r:2), (c:4; r:2), (c:5; r:2)),}
                                            {((c:3; r:3), (c:4; r:3), (c:5; r:3)),}
                                            ((c:3; r:4), (c:4; r:4), (c:5; r:4)),
                                            ((c:3; r:5), (c:4; r:5), (c:5; r:5)),
                                            ((c:3; r:6), (c:4; r:6), (c:5; r:6)),
                                            ((c:3; r:7), (c:4; r:7), (c:5; r:7)),
                                                                                    ((c:6; r:0), (c:7; r:0), (c:8; r:0)),
                                                                                    ((c:6; r:1), (c:7; r:1), (c:8; r:1)),
                                                                                    ((c:6; r:2), (c:7; r:2), (c:8; r:2)),
                                                                                    ((c:6; r:3), (c:7; r:3), (c:8; r:3)),
                                                                                    ((c:6; r:4), (c:7; r:4), (c:8; r:4)),
                                                                                    ((c:6; r:5), (c:7; r:5), (c:8; r:5)),
                                                                                    ((c:6; r:6), (c:7; r:6), (c:8; r:6)),
                                                                                    ((c:6; r:7), (c:7; r:7), (c:8; r:7)),
                                                                                                                          ((c:9; r:0), (c:10; r:0), (c:11; r:0)),
                                                                                                                          ((c:9; r:1), (c:10; r:1), (c:11; r:1)),
                                                                                                                          ((c:9; r:2), (c:10; r:2), (c:11; r:2)),
                                                                                                                          ((c:9; r:3), (c:10; r:3), (c:11; r:3)),
                                                                                                                          ((c:9; r:4), (c:10; r:4), (c:11; r:4)),
                                                                                                                          ((c:9; r:5), (c:10; r:5), (c:11; r:5)),
                                                                                                                          ((c:9; r:6), (c:10; r:6), (c:11; r:6)),
                                                                                                                          ((c:9; r:7), (c:10; r:7), (c:11; r:7))
     {$ENDREGION}
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
  end;

  TRawDataList = class(TStringList)
  private
    function GetMats(c, r: Integer): String;
    procedure SetMats(c, r: Integer; const Value: String);
  public
    function Cols(const r: Integer; var ACols: TStringList): Boolean;

    property Mats[c, r: Integer]: String read GetMats write SetMats;
  end;

  TBlockList = class
  private
    FGuard: TObject;
    FSeed: TRawDataList;
    FNsList: IDictionary<TMatrix, TBlock>;
    FStdList: IList<TBlock>;
    FSampleList: IList<TBlock>;
    FBlockIdx: TMatArray<Integer>;
    FBlockTypes: TMatArray<TBlockType>;
    FPointIdx: TMatArray<TPointIdx>;
    FColCount: Integer;
    procedure BuildSeed(const ASrc: TStringList);
    procedure BuildRawData;
    function GetBgColors(Col, Row: Integer): TColor;
    function GetCursor(Col, Row: Integer): TCursor;
    function GetFontStyles(Col, Row: Integer): TFontStyles;
    function GetMtrlTexts(Col, Row: Integer): String;
    function GetValues(Col, Row: Integer): String;
  public
    constructor Create;
    destructor Destroy; override;

    procedure Clear;
    procedure Paste(const ASrc: TStringList; const AColCount: Integer);
    procedure AssignDefault(const ACriteria: TCriteriaType);
    function ExtrctBlockFromNsList(const ACol, ARow: Integer; var ABlock: TBlock): Boolean; overload;
    function ExtractBlockFromNsList(const AMat: TMatrix; var ABlock: TBlock): Boolean; overload;

    property MtrlTexts[Col, Row: Integer]: String read GetMtrlTexts;
    property Values[Col, Row: Integer]: String read GetValues;
    property BgColors[Col, Row: Integer]: TColor read GetBgColors;
    property FontStyle[Col, Row: Integer]: TFontStyles read GetFontStyles;
    property Cursor[Col, Row: Integer]: TCursor read GetCursor;
  end;

implementation

uses
  m.rawdata,

  System.Math, System.StrUtils, Spring.SystemUtils, System.IniFiles, mStringListHelper, mExceptions
  ;

{ TMatrix }

constructor TMatrix.Create(const ACol, ARow: Integer);
begin
  c := ACol;
  r := ARow;
end;

{ TPoint }

constructor TPoint.CreateFromClipbrd(const ACol, ARow: Integer; const AValue: string);
begin
  FCol := ACol;
  FRow := ARow;
  FIdx := -1;
  FValue := AValue;
end;

constructor TPoint.CreateFromSectionValue(const ASrc: string);
var
  LSrc: TArray<String>;
begin
  LSrc := ASrc.Split(['|']);
  Assert(Length(LSrc) = 4, 'Invalid parameter: '+ ASrc);

  FCol := LSrc[0].ToInteger -1;
  FRow := LSrc[1].ToInteger -1;
  FIdx := LSrc[2].ToInteger;
  FIdx := IfThen(FIdx > 0, FIdx -1, FIdx);
  FValue := LSrc[3];
end;

function TPoint.GetAsDoubleValue: Double;
begin
  if IsValueOver then
    Result := 4.
  else
    Result := FValue.ToDouble
end;

function TPoint.GetHasValue: Boolean;
begin
  Result := not FValue.IsEmpty and not FValue.Equals('N/S');
end;

function TPoint.GetIsValueOver: Boolean;
begin
  Result := FValue.ToLower.Equals(SValueOver);
end;

function TPoint.GetMatrix: TMatrix;
begin
  Result := TMatrix.Create(Col, Row);
end;

function TPoint.GetValue: String;
begin
  Result := IfThen(FValue.IsEmpty, SValueNotSpecified, FValue);
end;

function TPoint.ToString: string;
begin
  Result := Format('%d|%d|%d|%s', [
    FCol +1,
    FRow +1,
    IfThen(FIdx > 0, FIdx +1, FIdx),
    IfThen(IsValueOver, '4', FValue)
  ]);
end;

{ TCriteriaTypeHelper }

class function TCriteriaTypeHelper.Create(const AValue: Integer): TCriteriaType;
begin
  case AValue of
    0: Result := ctNil_Antigen;
    1: Result := ctNil_Antigen_Mitogen;
    8: Result := ctStandard;
  else
    raise ELogical.Create('LogicalError, Can not handled Param: ' + AValue.ToString);
  end;
end;

function TCriteriaTypeHelper.DefaultBlock(const c, r: Integer): TBlockType;
begin
  if DefaultStdRange(c, r) then
    Result := btStandard
  else
    Result := btSample;
end;

function TCriteriaTypeHelper.DefaultIdx(const c, r: Integer): Integer;
begin
  Result := -1;
  case Self of
    ctNil_Antigen:
      if DefaultStdRange(c, r) then
        Result := IfThen(c > 5, 1)
      else
        Result := (c * 4) + (r div 2) - IfThen((c >= 5), 2 + IfThen(c > 5, 2));

    ctNil_Antigen_Mitogen:
      if DefaultStdRange(c, r) then
        Result := IfThen(c > 3, 1 + IfThen(c > 4, 1 + IfThen(c > 5, 1)))
      else
        Result := IfThen(c < 3, (c div 3), 4 + 8 * ((c div 3) -1)) + r;
  end;
  Assert(Result > -1, 'Exception occured when process the ' + ToString);
end;

function TCriteriaTypeHelper.DefaultStdRange(const c, r: Integer): Boolean;
begin
  Result := False;
  case Self of
    ctNil_Antigen: Result := TSample.IsM2StdRange(c, r);
    ctNil_Antigen_Mitogen: Result := TSample.IsM3StdRange(c, r);
  else
    Assert(False, 'Exception occured when process the ' + ToString);
  end;
end;

function TCriteriaTypeHelper.ToInteger: Integer;
begin
  Result := 0;
  case Self of
    ctNil_Antigen: Result := 0;
    ctNil_Antigen_Mitogen: Result := 1;
    ctStandard: Result := 8;
    else
      Assert(False, 'Handled code not exists!!');
  end;
end;

{ TBlock }

function TBlock.AddPoint(const AValue: TPoint): Integer;
begin
  FPoints := FPoints + [AValue];
  Result := PointCnt;
  case Result of
    1:
    begin
      FBlockType := btNotSpecified;
    end;

    2:
    begin
      FBlockType := btSample;
      FCriteriaType := ctNil_Antigen;
    end;

    3:
    begin
      FBlockType := btSample;
      FCriteriaType := ctNil_Antigen_Mitogen;
    end;

    4:
    begin
      FBlockType := btStandard;
      FCriteriaType := ctStandard;
    end;
  end;
end;

function TBlock.AddPointFromBlock(const ASeq: Integer; const ABlocks: TArray<TBlock>): Integer;
var
  LBlock: TBlock;
begin
  Result := -1;
  for LBlock in ABlocks do
    Assert(LBlock.PointCnt = 1, 'Invalid Paramerter: Cannot be processed the ABlocks which contains two or more points.');

  FRandomAssigned := Length(ABlocks) = 1;
  for LBlock in ABlocks do
    Result := AddPoint(LBlock.Points[0]);
  case FBlockType of
    btSample:
    begin
      FFixedID :=  ASeq;
      FId := Format('ID %d', [ASeq +1]);
    end;

    btStandard:
    begin
      FFixedID := 0;
      FId := Format('Std %d', [ASeq +1]);
    end;

    btNotSpecified: ;
  end;
end;

constructor TBlock.CreateFromClipbrd(const ACol, ARow: Integer; const AValue: string);
begin
  AddPoint(TPoint.CreateFromClipbrd(ACol, ARow, AValue));
  FCriteriaType := Nullable.Null;
  FFixedID := -1;
  FId := '';
end;

constructor TBlock.CreateFromPoint(const APoint: TPoint);
begin
  AddPoint(APoint);
end;

constructor TBlock.CreateFromSection(const CellIdx: Integer; ASec: TStrings; const AFreeSection: Boolean);
var
  LIdent: String;
  procedure AddIdentToPoint(const AIdentValue: String);
  begin
    AddPoint(TPoint.CreateFromSectionValue(AIdentValue));
  end;
begin
  Assert(ASec.KeyExists(SIdentPoints[0]), 'Invalid Paramerter: Point0 is not exists');
  Assert(ASec.KeyExists(SIdentBlockType), 'Invalid Paramerter: BlockType is not exists');

  for LIdent in SIdentPoints do
    if ASec.KeyExists(LIdent) then
      AddIdentToPoint(ASec.S[LIdent]);

  FRandomAssigned := PointCnt = 1;

  if ASec.KeyExists(SIdentCriteriaType) then
    FCriteriaType := TCriteriaType.Create(ASec.I[SIdentCriteriaType]);

  if ASec.KeyExists(SIdentFixedId) then
    FFixedID := ASec.I[SIdentFixedId];

  if ASec.KeyExists(SIdentId) then
    FID := ASec.S[SIdentId];

  if AFreeSection then
    ASec.Free;
end;

function TBlock.GetBgColors(Idx: TPointIdx): TColor;
begin
  Result := TColors.White;
  case BlockType of
    btSample:
      case FCriteriaType.Value of
        ctNil_Antigen:
          Result := IfThen(FRandomAssigned,
            TSample.clRandom[Min(FixedID -1, TSample.NM2MaxLen -1)],
            TSample.clM2[Min(FixedID -1, TSample.NM2MaxLen -1)]
          );

        ctNil_Antigen_Mitogen:
          Result := IfThen(FRandomAssigned,
            TSample.clRandom[Min(FixedID -1, TSample.NM2MaxLen -1)],
            TSample.clM3[Min(FixedID -1, TSample.NM2MaxLen -1)]
          );
      end;

    btStandard:
      Result := TSample.clMStd[Idx];

    btNotSpecified:
      Result := TColors.White;

    else
      Assert(False, 'Handled code not exists!!');
  end;
end;

function TBlock.GetCriteriaType: TCriteriaType;
begin
  Result := FCriteriaType.Value;
end;

function TBlock.GetCursor: TCursor;
begin
  Result := crDefault;
  case BlockType of
    btSample  ,
    btStandard:
      Result := crDrag;

    btNotSpecified:
      Result := crHandPoint;

    else
      Assert(False, 'Handled code not exists!!');
  end;
end;

function TBlock.GetFontStyles: TFontStyles;
begin
  case FBlockType of
    btSample      ,
    btNotSpecified:
      Result := [];

    btStandard:
      Result := [TFontStyle.fsBold, TFontStyle.fsUnderline];

    else
      Assert(False, 'Handled code not exists!!');
  end;
end;

function TBlock.GetLastPoint: TPoint;
begin
  Result := FPoints[PointCnt -1];
end;

function TBlock.GetMtrlTexts(Index: TPointIdx): String;
const
  SSamplePreFix: array[0..2] of Char = ('N', 'A', 'M');
var
  LPointIdx: Integer;
begin
  Result := '';
  LPointIdx := Points[Index].Idx;
  case LPointIdx of
    -1:
      Result := '';

     0, 1, 2:
      case BlockType of
        btSample  : Result := Format('%d%s', [LPointIdx +1, SSamplePreFix[LPointIdx]]);
        btStandard: Result := Format('S%d', [LPointIdx +1]);
      end;

     3:
      Result := Format('S%d', [LPointIdx +1]);
  end;
end;

function TBlock.GetPointCnt: Integer;
begin
  Result := Length(FPoints);
end;

function TBlock.GetPoints(Idx: TPointIdx): TPoint;
begin
  Result := FPoints[Idx];
end;

function TBlock.GetValues(Idx: TPointIdx): String;
begin
  Result := FPoints[Idx].Value;
end;

procedure TBlock.SetId(const Value: String);
begin
  if not Value.Equals(FId) and not Value.IsEmpty then
    FId := Value;
end;

function TBlock.ToNoSpecifiedBlocks: TArray<TBlock>;
var
  i: Integer;
begin
  Result := [];
  for i := 1 to Length(FPoints) -1 do
    Result := Result + [TBlock.CreateFromPoint(FPoints[i])];

  FCriteriaType := Nullable.Null;
  FFixedID := -1;
  FId := '';
  FBlockType := btNotSpecified;
end;

procedure TBlock.WriteBuf(const ASeq: Integer; const ABuf: TTextWriter);
const
  SSectionFmt = '[Cell%d]';
  SPointFmt = 'Point%d=%s';
  SCriterialTypeFmt = 'CriteriaType=%d';
  SFixedIDFmt = 'FixedID=%d';
  SIDFmt = 'ID=%s';
  SBlockTypeFmt = 'BlockType=%d';
var
  LPoint: TPoint;
begin
  Assert(Assigned(ABuf), 'ABuf is not assigned!!');

  ABuf.WriteLine(SSectionFmt, [ASeq]);
  for LPoint in FPoints do
    ABuf.WriteLine(SPointFmt, [LPoint.Idx +1, LPoint.ToString]);
  case BlockType of
    btSample  ,
    btStandard:
      begin
        ABuf.WriteLine(SCriterialTypeFmt, [CriteriaType.ToInteger]);
        ABuf.WriteLine(SFixedIDFmt, [IfThen(BlockType = btSample, FixedId)]);
        ABuf.WriteLine(SIDFmt, [Id]);
      end;

    btNotSpecified:;
  end;
  ABuf.WriteLine(SBlockTypeFmt, [BlockType.ToInteger]);
  if BlockType <> btNotSpecified then
    ABuf.WriteLine;
end;

{ TBlockTypeHelper }

class function TBlockTypeHelper.Create(const AValue: Integer): TBlockType;
begin
  Result := TEnum.Parse<TBlockType>(AValue)
end;

function TBlockTypeHelper.ToInteger: Integer;
begin
  Result := TEnum.GetValue<TBlockType>(Self)
end;

{ TBlockList }

procedure TBlockList.AssignDefault(const ACriteria: TCriteriaType);
var
  bi, ci: Integer;
  LBlock: TBlock;
  function SampleBlockMat(const ACrtrIdx, APointIdx: Integer): TMatrix;
  begin
    if ACriteria = ctNil_Antigen then
      Result := TSample.matM2[ACrtrIdx, APointIdx]
    else
      Result := TSample.matM3[ACrtrIdx, APointIdx];
  end;
  function StdBlockMat(const ACrtrIdx, APointIdx: Integer): TMatrix;
  begin
    if ACriteria = ctNil_Antigen then
      Result := TSample.matStdM2[ACrtrIdx, APointIdx]
    else
      Result := TSample.matStdM3[ACrtrIdx, APointIdx];
  end;
  function ExtractBlocks(const ACrtrIdx: Integer; const ABlockType: TBlockType): TArray<TBlock>;
  var
    pm: TMatrix;
    pi: Integer;
    LExtBlock: TBlock;
  begin
    Result := [];
    for pi := 1 to IfThen(ACriteria = ctNil_Antigen, 1, 2) do
    begin
      case ABlockType of
        btSample: pm := SampleBlockMat(ACrtrIdx, pi);
        btStandard: pm := StdBlockMat(ACrtrIdx, pi);
        else
          Continue;
      end;
      if ExtractBlockFromNsList(pm, LExtBlock) then
      begin
        Result := Result + [LExtBlock];
        FBlockIdx[pm.c, pm.r] := bi;
        FBlockTypes[pm.c, pm.r] := ABlockType;//btSample;
        FPointIdx[pm.c, pm.r] := pi;
      end;
    end;
  end;
  function ExtractNsBlockToList(const ADst: IList<TBlock>; const AMat: TMatrix; const ABlockType: TBlockType; var ABlock: TBlock): Boolean;
  begin
    Result := ExtractBlockFromNsList(AMat, ABlock);
    if Result then
    begin
      bi := ADst.Count;
      FBlockIdx[AMat.c, AMat.r] := bi;
      FBlockTypes[AMat.c, AMat.r] := ABlockType;
      FPointIdx[AMat.c, AMat.r] := ABlock.PointCnt -1;
    end;
  end;
begin
  FSampleList.Clear;
  for ci := 0 to IfThen(ACriteria = ctNil_Antigen, TSample.NM2Len, TSample.NM3Len) -1 do
    if ExtractNsBlockToList(FSampleList, SampleBlockMat(0, 0), btSample, LBlock) then
    begin
      LBlock.AddPointFromBlock(ci, ExtractBlocks(ci, btSample));
      FSampleList.Add(LBlock);
    end;

  FStdList.Clear;
  for ci := 0 to IfThen(ACriteria = ctNil_Antigen, 1, 2) do
    if ExtractNsBlockToList(FStdList, StdBlockMat(0, 0), btStandard, LBlock) then
    begin
      LBlock.AddPointFromBlock(bi, ExtractBlocks(ci, btStandard));
      FStdList.Add(LBlock);
    end;
end;

procedure TBlockList.BuildRawData;
var
  c, r: Integer;
  LCols: TStringList;
  mat: TMatrix;
begin
  FNsList.Clear;
  LCols := TStringList.Create;
  try
    LCols.StrictDelimiter := True;
    LCols.Delimiter := #9;
    for r := 0 to NRowCnt -1 do
    begin
      LCols.DelimitedText := FSeed[r];
      for c := 0 to LCols.Count -1 do
      begin
        mat := TMatrix.Create(c, r);
        FNsList.Add(mat, TBlock.CreateFromClipbrd(c, r, LCols[c]));
        // Block Index는 부여된 순서대로 할당되어야 한다.
        FBlockIdx[c, r] := FNsList.Count -1;
        FBlockTypes[c, r] := btNotSpecified;
        FPointIdx[c, r] := FNsList[mat].PointCnt -1;
      end;
    end;
  finally
    FreeAndNil(LCols);
  end;
end;

procedure TBlockList.BuildSeed(const ASrc: TStringList);
var
  LCols: TStringList;
  r: Integer;
begin
  LCols := TStringList.Create;
  try
    FSeed.Clear;
    FColCount := -1;
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
  finally
    FreeAndNil(LCols);
  end;
end;

procedure TBlockList.Clear;
begin
  FNsList.Clear;
  FStdList.Clear;
  FSampleList.Clear;
  FSeed.Clear;
end;

constructor TBlockList.Create;
begin
  FGuard := TObject.Create;
  FNsList := TCollections.CreateDictionary<TMatrix, TBlock>;
  FStdList := TCollections.CreateList<TBlock>;
  FSampleList := TCollections.CreateList<TBlock>;
  FSeed := TRawDataList.Create;
end;

destructor TBlockList.Destroy;
begin
  FreeAndNil(FSeed);
  FNsList := nil;
  FStdList := nil;
  FSampleList := nil;
  FreeAndNil(FGuard);

  inherited;
end;

function TBlockList.ExtractBlockFromNsList(const AMat: TMatrix; var ABlock: TBlock): Boolean;
begin
  Result := FNsList.ContainsKey(AMat);
  if Result then
  begin
    ABlock := FNsList[AMat];
    FNsList.Remove(AMat);
  end;
end;

function TBlockList.ExtrctBlockFromNsList(const ACol, ARow: Integer; var ABlock: TBlock): Boolean;
begin
  Result := ExtractBlockFromNsList(TMatrix.Create(ACol, ARow), ABlock);
end;

function TBlockList.GetBgColors(Col, Row: Integer): TColor;
var
  pi: Integer;
  mat: TMatrix;
begin
  Result := TColors.White;
  mat := TMatrix.Create(Col, Row);
  pi := FPointIdx[Col, Row];
  case FBlockTypes[Col, Row] of
    btSample:;
    btStandard:;
    btNotSpecified: Result := FNsList[mat].BgColors[pi];
  end;
end;

function TBlockList.GetCursor(Col, Row: Integer): TCursor;
var
  mat: TMatrix;
begin
  Result := crDefault;
  mat := TMatrix.Create(Col, Row);
  case FBlockTypes[Col, Row] of
    btSample:;
    btStandard:;
    btNotSpecified: Result := FNsList[mat].Cursor;
  end;
end;

function TBlockList.GetFontStyles(Col, Row: Integer): TFontStyles;
var
  mat: TMatrix;
begin
  Result := [];
  mat := TMatrix.Create(Col, Row);
  case FBlockTypes[Col, Row] of
    btSample:;
    btStandard:;
    btNotSpecified: Result := FNsList[mat].FontStyle;
  end;
end;

function TBlockList.GetMtrlTexts(Col, Row: Integer): String;
var
  pi: Integer;
  mat: TMatrix;
begin
  Result := '';
  mat := TMatrix.Create(Col, Row);
  pi := FPointIdx[Col, Row];
  case FBlockTypes[Col, Row] of
    btSample:;
    btStandard:;
    btNotSpecified: Result := FNsList[mat].MtrlTexts[pi];
  end;
end;

function TBlockList.GetValues(Col, Row: Integer): String;
var
  pi: Integer;
  mat: TMatrix;
begin
  Result := '';
  mat := TMatrix.Create(Col, Row);
  pi := FPointIdx[Col, Row];
  case FBlockTypes[Col, Row] of
    btSample:;
    btStandard:;
    btNotSpecified: Result := FNsList[mat].Values[pi];
  end;
end;

procedure TBlockList.Paste(const ASrc: TStringList; const AColCount: Integer);
begin
  TMonitor.Enter(FGuard);
  try
    BuildSeed(ASrc);
    BuildRawData;
  finally
    TMonitor.Exit(FGuard);
  end;
end;

{ TRawDataList }

function TRawDataList.Cols(const r: Integer; var ACols: TStringList): Boolean;
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

function TRawDataList.GetMats(c, r: Integer): String;
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

procedure TRawDataList.SetMats(c, r: Integer; const Value: String);
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

{ TSample }

class function TSample.IsM2StdRange(const c, r: Integer): Boolean;
begin
  Result := InRange(c, 5, 6) and InRange(r, 0, 3);
end;

class function TSample.IsM3StdRange(const c, r: Integer): Boolean;
begin
  Result := InRange(c, 3, 5) and InRange(r, 0, 3);
end;

end.
