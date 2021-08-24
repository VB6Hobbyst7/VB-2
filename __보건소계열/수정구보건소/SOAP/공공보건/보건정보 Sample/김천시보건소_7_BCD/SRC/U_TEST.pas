unit U_TEST;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TF_Test = class(TForm)
    Memo1: TMemo;
    Button1: TButton;
    Button2: TButton;
    edBcd: TEdit;
    Memo2: TMemo;
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
  F_Test: TF_Test;

implementation

uses U_Server, GlobalVar;

{$R *.dfm}

procedure TF_Test.Button1Click(Sender: TObject);
var
  S:string;
begin
  S:= WorkListCall;

  if S = ''  then
      Memo1.Text:= TGlobal.SvrError
  else
      Memo1.Text:= S;

end;

procedure TF_Test.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action:= caFree;
end;

procedure TF_Test.FormDestroy(Sender: TObject);
begin
  F_Test:= nil;
end;

procedure TF_Test.Button2Click(Sender: TObject);
var
  S:string;
begin
  if edBcd.Text = '' then begin
      ShowMessage('바코드를 입력하세요!');
      edBcd.SetFocus;
      exit;
  end;

  S:= OrderCall(edBcd.Text);;

  if S = ''  then
      Memo2.Text:= TGlobal.SvrError
  else
      Memo2.Text:= S;
end;

end.
