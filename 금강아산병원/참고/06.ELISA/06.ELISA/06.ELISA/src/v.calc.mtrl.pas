unit v.calc.mtrl;

interface

uses
  m.rawdata,

  mvw.vForm,

  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, AdvUtil, Vcl.ExtCtrls, Vcl.Grids, AdvObj, BaseGrid, AdvGrid, Vcl.StdCtrls,
  i18nCore, i18nLocalizer, Vcl.Imaging.pngimage, RzButton, RzRadChk, System.ImageList, Vcl.ImgList, PngImageList, RzTabs;

type
  TvCalcMtrl = class(TvForm)
    Translator: TTranslator;
    PngImageList1: TPngImageList;
    PageControl: TRzPageControl;
    TabM2: TRzTabSheet;
    TabM3: TRzTabSheet;
    GridM2: TAdvStringGrid;
    GridM3: TAdvStringGrid;
    procedure FormCreate(Sender: TObject);
    procedure FormResize(Sender: TObject);

    procedure GridGetCellColor(Sender: TObject; ARow, ACol: Integer; AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
  private
    FGrid: array[cmNil_Antigen .. cmNil_Antigen_Mitogen] of TAdvStringGrid;

    procedure InitGrid(const AGrd: TAdvStringGrid; const ACrtrMtrl: TCriteriaMaterial);
    function GridToString(const ACrtrMtrl: TCriteriaMaterial): String;
    function GetCells(c: String; r: Integer): String;
    function GetMtrlCnt: Integer;
    function GetCrtrMtrl: TCriteriaMaterial;
  public
    procedure Initailize;
    procedure Clear;
    procedure ExportToClilpboard;
    procedure ExportToCsvFile(const AFileName: String);

    property Cells[c: String; r: Integer]: String read GetCells;
    property MtrlCnt: Integer read GetMtrlCnt;
    property CrtrMtrl: TCriteriaMaterial read GetCrtrMtrl;
  end;

var
  vCalcMtrl: TvCalcMtrl;

implementation

{$R *.dfm}

uses
  svc,
  m.calc.mtrl,

  mAdvStringGridHelper, System.Math, System.StrUtils, System.UITypes, Vcl.Clipbrd, mUtils.Windows, System.DateUtils,
  mDateTimeHelper
  ;

function IfThen(const ACondition: Boolean; ATrue, AFalse: TRzTabSheet): TRzTabSheet; overload;
begin
  if ACondition then
    Exit(ATrue)
  else
    Exit(AFalse);
end;

function IfThen(const ACondition: Boolean; ATrue, AFalse: TAdvStringGrid): TAdvStringGrid; overload;
begin
  if ACondition then
    Exit(ATrue)
  else
    Exit(AFalse);
end;

{ TvMaterialResult }

procedure TvCalcMtrl.Clear;
begin
  //LabelResult.Caption := '';

//  TabM2.Enabled := False;
//  TabM3.Enabled := False;
  GridM2.Clear;
  GridM3.Clear;
end;

procedure TvCalcMtrl.ExportToClilpboard;
begin
  ClipBoard.Clear;
  ClipBoard.AsText := GridToString(CrtrMtrl);
end;

procedure TvCalcMtrl.ExportToCsvFile(const AFileName: String);
var
  LBuf: TStringStream;
begin
  LBuf := TStringStream.Create('', TEncoding.Unicode);
  try
    LBuf.SaveToFile(AFileName);
  finally
    FreeAndNil(LBuf);
  end;
end;

procedure TvCalcMtrl.FormCreate(Sender: TObject);
begin
  PageControl.ActivePageIndex := 0;

  FGrid[cmNil_Antigen] := GridM2;
  FGrid[cmNil_Antigen_Mitogen] := GridM3;
end;

procedure TvCalcMtrl.FormResize(Sender: TObject);
begin
  GridM2.AutoFitColumns(False);
  GridM3.AutoFitColumns(False);
end;

function TvCalcMtrl.GetCells(c: String; r: Integer): String;
var
  LCol: Integer;
  LGrd: TAdvStringGrid;
begin
  Result := '';
  LGrd := FGrid[CrtrMtrl];
  LCol := LGrd.Rows[0].IndexOf(c);
  if LCol = -1 then
    Exit;

  if (LCol < LGrd.ColCount) and (r < LGrd.RowCount)  then
    Result := FGrid[CrtrMtrl].Cells[LCol, r];
end;

function TvCalcMtrl.GetCrtrMtrl: TCriteriaMaterial;
begin
  Result := IfThen(PageControl.ActivePageIndex =  TabM2.PageIndex, cmNil_Antigen, cmNil_Antigen_Mitogen);
end;

function TvCalcMtrl.GetMtrlCnt: Integer;
begin
  Result := FGrid[CrtrMtrl].RowCount -1;
end;

procedure TvCalcMtrl.GridGetCellColor(Sender: TObject; ARow, ACol: Integer; AState: TGridDrawState;
  ABrush: TBrush; AFont: TFont);
var
  LGrd: TAdvStringGrid absolute Sender;
  LRet: String;
begin
  if gdFixed in AState then
    Exit;

  LRet := LGrd.Cells[LGrd.ColCount -1, ARow];
  if LRet.IsEmpty then
     Exit;

  ABrush.Style := bsSolid;
  AFont.Style := [];

    case LRet.Chars[0] of
      'N':
        if ACol = LGrd.ColCount -1 then
        begin
          AFont.Color := TColors.Black;
          //AFont.Style := [fsBold];
        end;

      'P':
      begin
        if ACol = LGrd.ColCount -1 then
        begin
          AFont.Color := TColors.Hotpink;
          AFont.Style := [fsBold];

        end;
        ABrush.Color := TColors.Ghostwhite;
      end;

      'I':
        //if ACol = LGrd.ColCount -1 then
          AFont.Color := TColors.Darkgray;

      else
        Exit;

    end;
end;

function TvCalcMtrl.GridToString(const ACrtrMtrl: TCriteriaMaterial): String;
var
  LBuf: TStringWriter;
  r: Integer;
  LGrd: TAdvStringGrid;
begin
  LGrd := FGrid[ACrtrMtrl];
  LBuf := TStringWriter.Create;
  try
    LBuf.WriteLine('Version: ' + ExeVersion);
    LBuf.WriteLine('Operator: ' + dataContainer.Properties.&Operator);
    LBuf.WriteLine('Kit Batch Number: ' + dataContainer.Properties.KitBatchNumber);
    LBuf.Write(LGrd.Cells[0, 0] +#9);
    LBuf.Write('Run Number' +#9);
    LBuf.Write('Run Date' +#9);
    LBuf.Write('Valid Test' +#9);
    LBuf.Write(LGrd.Cells[1, 0] +#9);
    LBuf.Write(LGrd.Cells[2, 0] +#9);
    LBuf.Write(LGrd.Cells[3, 0] +#9);
    LBuf.Write(LGrd.Cells[4, 0] +#9);
    LBuf.Write(LGrd.Cells[5, 0] +#9);
    LBuf.Write(LGrd.Cells[6, 0] +#9);
    LBuf.WriteLine;
    for r := 1 to LGrd.RowCount -1 do
    begin
      LBuf.Write(LGrd.Cells[0, r] +#9);
      LBuf.Write(dataContainer.Properties.RunNumber+#9);
      LBuf.Write(dataContainer.Properties.AsRunDateStr+#9);
      LBuf.Write(IfThen(stdCalc.Valid, 'Yes', 'No')+#9);
      LBuf.Write(LGrd.Floats[1, r].ToString + #9); //IUML[mNil, r];
      case ACrtrMtrl of
        cmNil_Antigen:
        begin
          LBuf.Write(LGrd.Cells[2, r]  +#9);//:= DoubleToCell(mtrlCalc.IUML[mTBAg, r]);
          LBuf.Write(LGrd.Cells[3, r]  +#9);//:= DoubleToCell(mtrlCalc.DeltaTBAgNil[r]);
          LBuf.Write(LGrd.Cells[4, r]  +#9);//:= mtrlCalc.ResultStrs[r];
        end;
        cmNil_Antigen_Mitogen:
        begin
          LBuf.Write(LGrd.Cells[2, r] +#9); //IUML[mTBAg, r]);
          LBuf.Write(LGrd.Cells[3, r] +#9); //IUML[mMitogen, r]);
          LBuf.Write(LGrd.Cells[4, r] +#9); //DeltaTBAgNil[r]);
          LBuf.Write(LGrd.Cells[5, r] +#9); //DeltaMtzNil[r]);
          LBuf.Write(LGrd.Cells[6, r] +#9); //ResultStrs[r];
        end;
      end;
      LBuf.WriteLine;
    end;
    Result := LBuf.ToString;
  finally
    FreeAndNil(LBuf);
  end;
end;

procedure TvCalcMtrl.Initailize;
begin
  mtrlCalc.Clear;
  if mtrlCalc.Execute(dataContainer.MtrlArray, stdCalc.Formula) then
  begin
    InitGrid(GridM2, cmNil_Antigen);
    InitGrid(GridM3, cmNil_Antigen_Mitogen);
  end;
  PageControl.ActivePageIndex := IfThen(TabM2.TabVisible, 0, 1);
end;

procedure TvCalcMtrl.InitGrid(const AGrd: TAdvStringGrid; const ACrtrMtrl: TCriteriaMaterial);
const
  SLabel: array[cmNil_Antigen .. cmNil_Antigen_Mitogen] of String = ('In-Tube(Nil, Antigen)', 'In-Tube(Nil, Antigen, Mitogen)');
var
  cmi, i: Integer;
  LPoints: TArray<TCellPoint>;
  LTab: TRzTabSheet;
  LCalc: TSamplesCalculator;
  function CalcMtrl(const AValue: String): String;
  var
    LValue, LLog10: Double;
  begin
    if not Double.TryParse(AValue, LValue) then
      Exit('N/S');

    LLog10 := Log10(LValue);
    Result := DoubleToString(Power(10, StdCalc.Formula.bX1 * LLog10 + StdCalc.Formula.cX0));
  end;
begin
  Assert(ACrtrMtrl in TCrtrMtrlSampleSet , 'Invalid parameter!!');
  LCalc := IfThen(ACrtrMtrl = cmNil_Antigen, MtrlCalc.M2, MtrlCalc.M3);
  LTab := IfThen(ACrtrMtrl = cmNil_Antigen, TabM2, TabM3);
  LTab.TabVisible := LCalc.Count > 0;
  if not LTab.TabVisible then
    Exit;

  AGrd.BeginUpdate;
  try
    //LabelResult.Caption := SLabel[mtrlCalc.CriteriaMaterial];
    AGrd.ColCount := IfThen(ACrtrMtrl = cmNil_Antigen, 5, 7);
    case ACrtrMtrl of
      cmNil_Antigen: AGrd.AssignCols(0, 0, ['Subject ID', 'Nil', 'TB Ag', 'TB Ag-Nil', 'Result']);
      cmNil_Antigen_Mitogen: AGrd.AssignCols(0, 0, ['Subject ID', 'Nil', 'TB Ag', 'Mitogen', 'TB Ag-Nil', 'Mitogen-Nil', 'Result']);
    end;
    AGrd.RowCount := LCalc.Count +1;
    for i := 0 to LCalc.Count -1 do
    begin
      cmi := LCalc.SrcIdxs[i];
      AGrd.Cells[0, i +1] := dataContainer.IDs[cmi];
      if LCalc.Exists(mNil, i) then
      begin
        AGrd.Cells[1, i +1] := DoubleToString(LCalc.SrcIumls[mNil, i]);
        case ACrtrMtrl of
          cmNil_Antigen:
          begin
            AGrd.Cells[2, i +1] := DoubleToString(LCalc.SrcIumls[mTBAg, i]);
            AGrd.Cells[3, i +1] := DoubleToString(LCalc.IumlDeltaTBAg[i]);
            AGrd.Cells[4, i +1] := LCalc.ResultTexts[i];
          end;
          cmNil_Antigen_Mitogen:
          begin
            AGrd.Cells[2, i +1] := DoubleToString(LCalc.SrcIumls[mTBAg, i]);
            AGrd.Cells[3, i +1] := DoubleToString(LCalc.SrcIumls[mMitogen, i]);
            AGrd.Cells[4, i +1] := DoubleToString(LCalc.IumlDeltaTBAg[i]);
            AGrd.Cells[5, i +1] := DoubleToString(LCalc.IumlDeltaMtz[i]);
            AGrd.Cells[6, i +1] := LCalc.ResultTexts[i];
          end;
        end;
      end
      else
      begin
        LPoints := dataContainer.PointsByIds[cmi];
        AGrd.Cells[1, i +1] := CalcMtrl(LPoints[0].Value);
        case ACrtrMtrl of
          cmNil_Antigen:
          begin
            AGrd.Cells[2, i +1] := CalcMtrl(LPoints[1].Value);
            AGrd.Cells[3, i +1] := '-';
            AGrd.Cells[4, i +1] := 'Data Missing';
          end;
          cmNil_Antigen_Mitogen:
          begin
            AGrd.Cells[2, i +1] := CalcMtrl(LPoints[1].Value);
            AGrd.Cells[3, i +1] := CalcMtrl(LPoints[2].Value);
            AGrd.Cells[4, i +1] := '-';
            AGrd.Cells[5, i +1] := '-';
            AGrd.Cells[6, i +1] := 'Data Missing';
          end;
        end;
      end;
    end;
    AGrd.AutoFitColumns;
    AGrd.AutoGrowCol(AGrd.ColCount -1);
  finally
    AGrd.EndUpdate;
  end;
end;

end.
