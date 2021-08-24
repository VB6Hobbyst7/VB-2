unit U_TEST;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TF_Test = class(TForm)
    Memo1: TMemo;
    Button1: TButton;
    procedure Button1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  F_Test: TF_Test;

implementation

uses U_Server;

{$R *.dfm}

procedure TF_Test.Button1Click(Sender: TObject);
begin
  Memo1.Text:= WorkListCall;
end;

procedure TF_Test.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action:= caFree;
end;

procedure TF_Test.FormDestroy(Sender: TObject);
begin
  F_Test:= nil;
end;

end.
