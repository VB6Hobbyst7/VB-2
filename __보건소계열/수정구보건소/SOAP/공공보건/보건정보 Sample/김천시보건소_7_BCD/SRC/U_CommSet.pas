unit U_CommSet;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls;

type
  TF_CommSet = class(TForm)
    GroupBox1: TGroupBox;
    cmbPortNum: TComboBox;
    cmbBaudrate: TComboBox;
    cmbDatabits: TComboBox;
    cmbStopbits: TComboBox;
    cmbParity: TComboBox;
    cmbHand: TComboBox;
    Panel2: TPanel;
    Panel1: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    cbxDTR: TCheckBox;
    cbxRTS: TCheckBox;
    Button1: TButton;
    Button2: TButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
    procedure SetPortNum(iPort:integer);
    procedure SetBoudRate(sBaud: string);
    procedure SetDataBit(sDb: string);
    procedure SetStopBit(sSb: string);
    procedure SetParity(sPt:string);
    procedure SetHShake(iHs:integer);
    procedure SetDtr(bDtr:boolean);
    procedure SetRts(bRts:boolean);
  public
    { Public declarations }
    TCommPort:TObject;
    procedure Load;
    procedure Save;
  end;


  type
    EPortTypeError = Class(EAbort)
      Constructor Create;
    End;
var
  F_CommSet: TF_CommSet;

implementation

uses StringLib, GlobalVar, MSCommLib_TLB, CPort;

{$R *.dfm}

{ TF_MSCommSet }

procedure TF_CommSet.Button1Click(Sender: TObject);
begin
  if MessageDlg('장비와의 통신을 종료후 재시작 합니다. '#13#10' 계속 하시겠습니까?', mtWarning, mbOKCancel, 1) = mrOk then
  begin
      Save;
      ModalResult:= mrOk;
  end;

end;

procedure TF_CommSet.Button2Click(Sender: TObject);
begin
  ModalResult:= mrCancel;
end;

procedure TF_CommSet.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action:= caFree;
end;

procedure TF_CommSet.FormDestroy(Sender: TObject);
begin
  F_CommSet:= nil;
end;

procedure TF_CommSet.FormShow(Sender: TObject);
begin
  Load;
end;

procedure TF_CommSet.Load;
var
  SetStr:string;
begin
  with TMsComm(TCommPort) do begin
      SetStr:= Settings+',';
      SetPortNum(CommPort);
      SetBoudRate(Trim(TokenStr(Settings,',',1)));
      SetParity(Trim(TokenStr(Settings,',',2)));
      SetDataBit(Trim(TokenStr(Settings,',',3)));
      SetStopBit(Trim(TokenStr(Settings,',',4)));
      SetHShake(Handshaking);
      SetDtr(DTREnable);
      SetRts(RTSEnable);
  end;
end;

procedure TF_CommSet.Save;
var
  SetStr:string;
begin
  if TCommPort Is TMSComm then begin
      with TMsComm(TCommPort) do begin
          if PortOpen then
              PortOpen:= False;

          RTSEnable:= cbxRTS.Checked;
          dtrEnable:= cbxDTR.Checked;
          CommPort := StrToInt(cmbPortNum.Text);
          SetStr:=  cmbBaudrate.Text + ',';
          { Parity = E 짝수, M 표시, N  (기본값)없음, O 홀수, S 공간}
          SetStr:= SetStr + Copy(cmbParity.Text,1,1)+ ',';
          SetStr:= SetStr + cmbDatabits.Text + ',';
          SetStr:= SetStr + cmbStopbits.Text;

          Settings:= SetStr;
          Handshaking:= cmbHand.ItemIndex;

          PortOpen:= True;

          TGlobal.ComPortSet.PortNum := CommPort;
          TGlobal.ComPortSet.BaudRate:= cmbBaudrate.Text;
          TGlobal.ComPortSet.Parity  := cmbParity.Text;
          TGlobal.ComPortSet.DataBit := cmbDatabits.Text;
          TGlobal.ComPortSet.StopBit := cmbStopbits.Text;
          TGlobal.ComPortSet.HandShake:= Handshaking;
          TGlobal.ComPortSet.Rts:= cbxRTS.Checked;
          TGlobal.ComPortSet.Dtr:= cbxDTR.Checked;
          TGlobal.ComPortIniSave;
          
      end;
  end
  else
      raise EPortTypeError.Create;

end;

procedure TF_CommSet.SetBoudRate(sBaud: string);
begin
  cmbBaudrate.ItemIndex:= cmbBaudrate.Items.IndexOf(sBaud);
end;

procedure TF_CommSet.SetDataBit(sDb: string);
begin
  cmbDatabits.ItemIndex:= cmbDatabits.Items.IndexOf(sDb);
end;

procedure TF_CommSet.SetDtr(bDtr: boolean);
begin
  cbxDTR.Checked:= bDtr;
end;

procedure TF_CommSet.SetHShake(iHs:integer);
begin
  cmbHand.ItemIndex:= iHS;
end;

procedure TF_CommSet.SetParity(sPt: string);
var
  i:integer;
begin
  for i := 0 to cmbParity.Items.Count - 1 do begin
      if LowerCase(Copy(cmbParity.Items.Strings[i],1,1)) = sPt then
          cmbParity.ItemIndex:= i;
  end;
end;

procedure TF_CommSet.SetPortNum(iPort: integer);
begin
  cmbPortNum.ItemIndex:= cmbPortNum.Items.IndexOf(IntToStr(iPort));
end;

procedure TF_CommSet.SetRts(bRts: boolean);
begin
  cbxRTS.Checked:= bRts;
end;

procedure TF_CommSet.SetStopBit(sSb: string);
begin
  cmbStopbits.ItemIndex:= cmbStopbits.Items.IndexOf(sSb);
end;

{ EPortTypeError }

constructor EPortTypeError.Create;
begin
  ShowMessage('맞지않는 포트 타입입니다!');
end;

end.
