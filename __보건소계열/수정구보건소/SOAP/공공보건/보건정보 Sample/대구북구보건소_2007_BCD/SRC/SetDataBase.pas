unit SetDataBase;

interface

uses ADODB, Forms, SysUtils, Dialogs;

type
  TDbConnection = Class(TObject)
    constructor Create;
    Destructor Destroy; override;
    private
      ConLocal:TAdoConnection;
      function SetConnectionString:boolean;
      function TestConnection:boolean;
    public
      procedure ChangeDbConnection(RCapDbConString:string);
      function SetConnection:boolean;
      property LocalCon: TAdoConnection read ConLocal;
      procedure AllDisconect;
  end;

  TQueryInfo = Class(TObject)
  private
    function ExcuteProcess(var TQryEx:TAdoQuery):boolean;
    function SelectProcess(var TQryEx:TAdoQuery):integer;
  public
    SqlCmd : string;
    RCount:integer;
    function LocalSelect(var TQryEx:TAdoQuery):integer;
    function LocalExcute:boolean;
    procedure AddSql(cStr:string);
    procedure Clear;

  end;

var
  TConnection: TDbConnection;

implementation

uses GlobalVar;

{ TDbConnection }

procedure TDbConnection.AllDisconect;
begin
  conLocal.Connected:= False;
end;

procedure TDbConnection.ChangeDbConnection(RCapDbConString: string);
begin
  conLocal.Connected:=False;
end;

constructor TDbConnection.create;
begin
  inherited;
  ConLocal:= TAdoConnection.Create(Application);
  ConLocal.LoginPrompt:= False;
  SetConnectionString;
end;

destructor TDbConnection.Destroy;
begin
  //conLocal.Connected:=False;
  inherited;
end;


function TDbConnection.SetConnection: boolean;
begin
  Result:= False;

  if Not SetConnectionString then
  begin
      ShowMessage(TGlobal.LastMsg);
      exit;
  end;

  Result:= True;
end;

function TDbConnection.SetConnectionString: boolean;
begin
  Result:= False;

  conLocal.Connected:=False;
  ConLocal.ConnectionString:= 'Provider=Microsoft.Jet.OLEDB.4.0;Password="";Data Source='+TGlobal.AppPath+
                              'SANSOFT.MDB;Persist Security Info=True ';

  Result:= TestConnection;

end;

function TDbConnection.TestConnection: boolean;
begin
  Result:= False;

  try
      Try
          ConLocal.Connected:= True;
      Except
          on e:Exception do
          begin
              TGlobal.LastMsg:= e.Message;
              exit;
          end;
      end;
      Result:= True;
  finally
      if ConLocal.Connected then
          ConLocal.Connected:= False;
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
              TGlobal.LogMsg:= e.Message+' [Sql:'+SqlCmd+']';
          end;
      end;
  end;

  if nCount > 0 then
          Result:= True
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
              TGlobal.LogMsg:= e.Message+Err1;
              exit;
          end;
      end;
  end;

  Result:= TQryEx.RecordCount;

end;

end.
