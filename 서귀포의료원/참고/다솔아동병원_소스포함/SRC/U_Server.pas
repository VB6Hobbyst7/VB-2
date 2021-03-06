unit U_Server;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, DB, Grids, DBGrids, ADODB, ComCtrls, ExtCtrls;

type
  TF_Server = class(TForm)
    Panel1: TPanel;
    StatusBar1: TStatusBar;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    mmSQL: TMemo;
    ADOQuery1: TADOQuery;
    DBGrid1: TDBGrid;
    DataSource1: TDataSource;
    Button1: TButton;
    Button2: TButton;
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  F_Server: TF_Server;

implementation

uses SetDataBase;

{$R *.dfm}

procedure TF_Server.FormCreate(Sender: TObject);
begin
  ADOQuery1.Connection:= TConnection.hospCon;
end;

procedure TF_Server.Button1Click(Sender: TObject);
begin
  with ADOQuery1 do begin
      Close;
      if mmSQL.SelText <> '' then
          SQL.Text:= mmSQL.SelText
      else
          SQL.Text:= mmSQL.Text;

      Active:= True;
  end;
end;

procedure TF_Server.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action:= caFree;
end;

procedure TF_Server.FormDestroy(Sender: TObject);
begin
  F_Server:= nil;
end;

procedure TF_Server.Button2Click(Sender: TObject);
begin
  with ADOQuery1 do begin
      Close;
      if mmSQL.SelText <> '' then
          SQL.Text:= mmSQL.SelText
      else
          SQL.Text:= mmSQL.Text;

      ExecSQL;
  end;
end;

end.
