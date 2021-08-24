unit SetDataBase;

interface

uses DBTables, ADODB, Forms, SysUtils, Dialogs;

const
  SvrConnection = True;
  SvrTEST = False;

type
  TDbConnection = Class(TObject)
    constructor Create;
    Destructor Destroy; override;
    private
      ConLocal:TAdoConnection;
      ConHosp:TAdoConnection;
      function SetConnectionString:boolean;
      function TestConnection:boolean;
    public
      function SetConnection:boolean;
      property LocalCon: TAdoConnection read ConLocal;
      property hospCon: TADOConnection read ConHosp;
  end;

  TQueryInfo = Class(TObject)
  private
    function ExcuteProcess(var TQryEx:TAdoQuery):boolean;  overload;
    function SelectProcess(var TQryEx:TAdoQuery):integer;  overload;
    function ExcuteProcess(var TQryEx:TQuery):boolean;  overload;
    function SelectProcess(var TQryEx:TQuery):integer;  overload;

  public
    SqlCmd : string;
    RCount:integer;
    function LocalSelect(var TQryEx:TAdoQuery):integer;
    function HospSelect(var TQryEx:TAdoQuery):integer;
    function LocalExcute:boolean;
    function HospExcute:boolean;
    procedure AddSql(cStr:string);
    procedure Clear;

  end;

var
  TConnection: TDbConnection;

implementation

uses GlobalVar;

{ TDbConnection }

constructor TDbConnection.create;
begin
  inherited;
  ConLocal:= TAdoConnection.Create(Application);
  ConLocal.LoginPrompt:= False;

  ConHosp:= TADOConnection.Create(Application);
  ConHosp.LoginPrompt:= False;

  SetConnectionString;

end;

destructor TDbConnection.Destroy;
begin
  inherited;
end;


function TDbConnection.SetConnection: boolean;
begin
  Result:= False;

  SetConnectionString;

  Result:= True;
end;

function TDbConnection.SetConnectionString: boolean;
begin
  Result:= False;

  conLocal.Connected:=False;
  ConLocal.ConnectionString:= 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source='+TGlobal.AppPath+
                              'SANSOFT.MDB;Persist Security Info=True ';
  ConHosp.Connected:=False;
  ConHosp.ConnectionString:= 'Provider=MSDAORA.1;Password=EON_SPP;User ID=EON_SPP;Data Source=EES;Persist Security Info=True';

  ConHosp.LoginPrompt:= False;
  ConHosp.KeepConnection:= True;

  Result:= TestConnection;
end;

function TDbConnection.TestConnection: boolean;
var
  conSucc:boolean;
begin
  Result:= False;
  conSucc:= True;

  try
      Try
          ConLocal.Connected:= True;
      Except
          on e:Exception do begin
              TGlobal.LastMsg:= e.Message;
              ShowMessage(e.Message);
              conSucc:= False;
          end;
      end;

      if SvrConnection = True then begin
          try
              ConHosp.Connected:= True;
          except
              on e:Exception do begin
                  TGlobal.LastMsg:= e.Message;
                  ShowMessage(e.Message);
                  conSucc:= False;
              end;
          end;
      end
      else begin
          ShowMessage('Local 테스트중 입니다!');
      end;

      Result:= conSucc;

  finally
      if ConLocal.Connected = True then
          ConLocal.Connected:= False;
      if ConHosp.Connected = True then
          ConHosp.Connected:= False;
  end;
end;

{ TQueryInfo }

procedure TQueryInfo.AddSql(cStr: string);
begin
  if Trim(SqlCmd) = '' then
      SqlCmd:= cStr
  else
      SqlCmd:= SqlCmd + #13#10 + cStr;
end;

procedure TQueryInfo.Clear;
begin
  SqlCmd:='';
end;

function TQueryInfo.ExcuteProcess(var TQryEx: TAdoQuery): boolean;
var
  nCount:integer;
begin
  Result:= False;
  nCount:=0;

  with TQryEx do
  begin
      Close;
      SQL.Text:= SqlCmd;
      try
          nCount:= TQryEx.ExecSQL;
      except
          on  e: Exception do
          begin
              TGlobal.LastMsg:= e.Message;
              TGlobal.ErrMsg:= e.Message+' [Sql:'+SqlCmd+']';
              exit;
          end;
      end;
  end;

  Result:= True
end;

function TQueryInfo.ExcuteProcess(var TQryEx: TQuery): boolean;
begin
  Result:= False;

  with TQryEx do begin
      Close;
      SQL.Text:= SqlCmd;
      try
          TQryEx.ExecSQL
      except
          on  e: Exception do begin
              TGlobal.LastMsg:= e.Message;
              TGlobal.ErrMsg:= e.Message+' [Sql:'+SqlCmd+']';
              exit;
          end;
      end;
  end;

  Result:= True

end;

function TQueryInfo.hospExcute: boolean;
var
  TQry: TADOQuery;
begin
  Result:= False;

  if SvrConnection = False then begin
      ShowMessage('Local 테스트중입니다!');
      exit;
  end;

  TQry:= TADOQuery.Create(Application);
  try
      TQry.Connection:= TConnection.hospCon;
      Result:= ExcuteProcess(TQry);

  finally
      TQry.Free;
  end;

end;

function TQueryInfo.HospSelect(var TQryEx: TAdoQuery): integer;
begin
  TQryEx.Close;
  TQryEx.Connection:= TConnection.ConHosp;
  Result:= SelectProcess(TQryEx);
end;

function TQueryInfo.LocalExcute: boolean;
var
  TQry: TAdoQuery;
begin
  Result:= False;

  TQry:= TADOQuery.Create(Application);
  try
      TQry.Connection:= TConnection.LocalCon;
      Result:= ExcuteProcess(TQry);

  finally
      TQry.Free;
  end;

end;

function TQueryInfo.LocalSelect(var TQryEx: TAdoQuery): integer;
begin
  TQryEx.Close;
  TQryEx.Connection:= TConnection.LocalCon;
  Result:= SelectProcess(TQryEx);
end;

function TQueryInfo.SelectProcess(var TQryEx: TAdoQuery): integer;
var
  Err1:string;
begin
  Result:= 0;

  with TQryEx do
  begin
      Close;
      SQL.Text:= SqlCmd;
      try
          Open;
      except
          on  e: Exception do
          begin
              TGlobal.LastMsg:= e.Message;
              Err1:=#13#10'['+SqlCmd+']';
              TGlobal.ErrMsg:= e.Message+Err1;
              exit;
          end;
      end;
  end;

  Result:= TQryEx.RecordCount;

end;

function TQueryInfo.SelectProcess(var TQryEx: TQuery): integer;
var
  Err1:string;
begin
  Result:= 0;

  with TQryEx do begin
      Close;
      SQL.Text:= SqlCmd;
      try
          Open;
      except
          on  e: Exception do begin
              TGlobal.LastMsg:= e.Message;
              Err1:=#13#10'['+SqlCmd+']';
              TGlobal.ErrMsg:= e.Message+Err1;
              exit;
          end;
      end;
  end;

  Result:= TQryEx.RecordCount;

end;

end.
