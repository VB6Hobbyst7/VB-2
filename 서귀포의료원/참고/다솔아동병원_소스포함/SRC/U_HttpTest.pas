unit U_HttpTest;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TF_ENCODE = class(TForm)
    Memo1: TMemo;
    Button1: TButton;
    Memo2: TMemo;
    Label1: TLabel;
    Label2: TLabel;
    procedure Button1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  F_ENCODE: TF_ENCODE;

implementation

uses HTTPApp;

{$R *.dfm}

procedure TF_ENCODE.Button1Click(Sender: TObject);
begin
  Memo2.Text:= HTTPDecode(Memo1.Text);
end;

procedure TF_ENCODE.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action:= caFree;
end;

procedure TF_ENCODE.FormDestroy(Sender: TObject);
begin
  F_Encode:= nil;
end;

end.
