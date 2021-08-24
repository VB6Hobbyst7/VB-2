unit U_IFClass;

interface

uses SysUtils;

type
  TIfMaster = Class(TObject)
    FExamDate,
    FInstTime,
    FRcvTime,
    FExamSeq,
    FBarCode,
    FRack, FPos,
    FPID,
    FPNM,
    FAcptDt,
    FAcptNo,
    FWorkNo,
    FAge,FSex,
    FLoc,FFlag,
    FRMK,
    FErYN,
    FICode,
    FResult,
    //QC
    FTYP,FLotLev,FLotMin,FLotMax,FLotMean,FLotSD,
    //
    FReMark,
    FUserID,
    FIfCode,
    FUpCode,
    FAbbr,
    FExamPanel,
    FOrdCode,
    FOrdName,
    FSUB,
    FOrdDate,
    FIpDate,
    FOrdNo,
    FOrdSeq,
    FANO,
    FADT,
    FLAB,
    FIO,
    FSLP,
    FExamCode,
    FQCYN,
    FPOCTYN,
    FLotNo,
    FDept,
    FUpState,
    SvrMsg,
    FOrdState:string;
    FRefLow:double;
    FRefHigh:double;
    FOrdCnt:integer;
    vOrder, vAbbr, vOrdList, vUpCode, vOrdCdList, vANO:Variant;
    OrdCnt:integer;
    //제일병원
    PRSNORNO:integer;
    PRSNORSQ:integer;
    //충북대병원
    SUBCODESEQ , CodeUnit, REFV:string;

    IsDownCodeOK:boolean;

    constructor Create;
  private
    FReceiptNo:integer;
    function GetANo: string;
    function GetANoStr: string;
    procedure SetANoStr(const Value: string);
    procedure SetReceiptNo(const Value: integer);
  public
    property ReceiptNo:integer read FReceiptNo write SetReceiptNo;
    property ANOStr:string read GetANoStr write SetANoStr;
    procedure Clear;
  end;

function CheckLowHigh(sLow,sHigh,sResult:string):string;

implementation

uses U_CodeInfo, StringLib, SetDataBase, DB;

function CheckLowHigh(sLow,sHigh,sResult:string):string;
var
  dMin,dMax,dVal:double;
begin
  Result:= '';

  dVal:= StrToFloatDef(sResult,-100);
  if dVal < -99 then exit;

  dMin:= StrToFloatDef(sLow, -100);
  dMax:= StrToFloatDef(sHigh, -100);

  if (dMin < -99) or (dMax < -99) then
      exit;

  if dVal < dMin then
      Result:= 'L'
  else
  if dVal > dMax then
      Result:= 'H';
end;

{ TIfMaster }

procedure TIfMaster.Clear;
begin
  FExamDate  :='';
  FInstTime  :='';
  FRcvTime   :='';
  FExamSeq   :='';
  FBarCode   :='';
  FRack      :='';
  FPos       :='';
  FPID       :='';
  FPNM       :='';
  FAcptDt    :='';
  FAcptNo    :='';
  FWorkNo    :='';
  FAge       :='';
  FSex       :='';
  FLoc       :='';
  FFlag      :='';
  FRMK       :='';
  FErYN      :='';
  FICode     :='';
  FResult    :='';
  FReMark    :='';
  FUserID    :='';
  FIfCode    :='';
  FAbbr      :='';
  FExamPanel :='';
  FOrdCode   :='';
  FSUB       :='';
  FOrdDate   :='';
  FIpDate    :='';
  FOrdNo     :='';
  FOrdSeq    :='';
  FANO       :='';
  FADT       :='';
  FLAB       :='';
  FIO        :='';
  FSLP       :='';
  FExamCode  :='';
  FQCYN      :='';
  FPOCTYN    :='';
  FLotNo     :='';
  FDept      :='';
  FUpState   :='';
  SvrMsg     :='';
  FOrdState  :='';
  FRefLow:=0;
  FRefHigh:=0;
  FOrdCnt:=0;
  OrdCnt:=0;
  FUpState:= 'N';
  FOrdState:= 'N';
  FQCYN:='N';
end;

constructor TIfMaster.Create;
begin
  FUpState:= 'N';
  FOrdState:= 'N';
  FQCYN:='N';
end;

function TIfMaster.GetANo: string;
begin
  Result:= IntToStr(FReceiptNo);
end;

function TIfMaster.GetANoStr: string;
begin
  Result:= IntToStr(FReceiptNo);
end;

procedure TIfMaster.SetANoStr(const Value: string);
begin
  FReceiptNo:= StrToIntDef(Value, 0);
end;

procedure TIfMaster.SetReceiptNo(const Value: integer);
begin
  FReceiptNo := Value;
end;

end.
