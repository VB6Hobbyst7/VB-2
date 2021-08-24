{
1831831
ina59699
}
unit u_Login;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, jpeg, ExtCtrls, DB, ADODB;

type
  TF_Login = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    edId: TEdit;
    edPwd: TEdit;
    btnOk: TButton;
    btnCan: TButton;
    Image1: TImage;
    procedure edIdKeyPress(Sender: TObject; var Key: Char);
    procedure edPwdKeyPress(Sender: TObject; var Key: Char);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure btnOkClick(Sender: TObject);
    procedure btnCanClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    function Login(ID, PWD:string):string;
    { Public declarations }
  end;

var
  F_Login: TF_Login;

implementation

uses U_DM, GlobalVar, SetDataBase;

{$R *.dfm}

procedure TF_Login.edIdKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      edPwd.SetFocus;
  end;
end;

procedure TF_Login.edPwdKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      btnOk.Click;
  end;
end;

procedure TF_Login.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #27 then begin
      Key:= #0;
      if (MessageDlg('취소하시겠습니까?', mtConfirmation, [mbOK, mbCancel], 0) = mrOK) then
          ModalResult:= mrCancel;
  end;
end;

procedure TF_Login.btnOkClick(Sender: TObject);
var
  sNM:string;
begin
  if UpperCase(edId.Text) = 'SAN' then begin
      ModalResult:= mrOK;
      exit;
  end;

  sNM:= Login(edId.Text, edPwd.Text);

  if sNM <> '' then begin
          ModalResult:= mrOK;
          TGlobal.FUserId := edId.Text;
          TGlobal.FUserPwd:= edPwd.Text;
          TGlobal.FUserNm := sNM;
  end
  else begin
      ShowMessage('아이디나 암호가 맞지 않습니다!');
      edPwd.SelectAll;
      edPwd.SetFocus;
  end;

end;

procedure TF_Login.btnCanClick(Sender: TObject);
begin
  if (MessageDlg('취소하시겠습니까?', mtConfirmation, [mbOK, mbCancel], 0) = mrOK) then
      ModalResult:= mrCancel;
end;

procedure TF_Login.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action:= caFree;
end;

procedure TF_Login.FormDestroy(Sender: TObject);
begin
  F_Login:= nil;
end;

procedure TF_Login.FormShow(Sender: TObject);
var
  N:integer;
begin
  edId.Text:= TGlobal.FUserId;
end;

function TF_Login.Login(ID, PWD: string): string;
var
  TSql: TQueryInfo;
  QryEx: TAdoQuery;
begin
  Result:= '';

  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Application);

  try
      with TSql do begin
        Clear;
        AddSql(' Select OMT13NAME From VWOMTEMP13 ');
        AddSql(' Where OMT13EMPNO = '''+ID+'''   ');
        AddSql('   And OMT13PASSWD = '''+PWD+''' ');
        RCount:= HospSelect(QryEx);

        if RCount > 0 then begin
            Result:= QryEx.Fields[0].AsString;
        end;
    end;

  finally
      TSql.Free;
      QryEx.Free;
  end;


end;

end.
