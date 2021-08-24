unit U_MsCommSet_Class;

interface

type
  TPortNum = (Com1=1, Com2, Com3, Com4, Com5, Com6, Com7, Com8);
  TDataBit = (Db5, Db6, Db7, Db8);
  TStopBit = (sb10, sb15, sb20);
  THandShake = (hsNone, hsXonXoff, hsRts, hsRtsXonXoff);

type
  TMSComm_Set = Class(TObject)
  private
    FPortNum: TPortNum;
    FSetting: string;
    FDtr:boolean;
    FRts:boolean;
    FHandShake: THandShake;
    FDataBit: TDataBit;
    FStopBit: TStopBit;
    procedure SetSetting(SetStr:string);
  public
    constructor Create;
    procedure Save;
    procedure Load;
    property PortNum:TPortNum read FPortNum write FPortNum;
    property Settings: string read FSetting write SetSetting;
    property Dtr:boolean read FDtr write FDtr;
    property Rts:boolean read FRts write FRts;
    property Databit:TDataBit read FDataBit write FDataBit;
    property StopBit:TStopBit read FStopBit write FStopBit;
    property HandShake:THandShake read FHandShake write FHandShake;
  end;

type
  TPortSet = Class(TObject)
    FType: CommType;
    FProtNum:integer;
    FBoudrate:string;
    FDataBit:integer;
    FStopBit:integer;
  End;

implementation

uses MSCommLib_TLB, CPort;

{ TMSComm_Set }

constructor TMSComm_Set.Create;
begin
  Load;
  FPortNum:= Com1;
  FSetting:= '';
  FDtr:= False;
  FRts:= False;
  FHandShake= hsNone;
  FDataBit:= db8;
  FStopBit:= sb10;
end;

procedure TMSComm_Set.Load;
begin

end;

procedure TMSComm_Set.Save;
begin

end;

procedure TMSComm_Set.SetSetting(SetStr: string);
begin

end;


end.
