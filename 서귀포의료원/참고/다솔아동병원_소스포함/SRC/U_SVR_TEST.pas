unit U_SVR_TEST;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls;

type
  TF_TEST = class(TForm)
    edBCD: TEdit;
    Panel1: TPanel;
    mmWork: TMemo;
    btnWSel: TButton;
    edID: TEdit;
    edPWD: TEdit;
    btnLSel: TButton;
    Panel2: TPanel;
    mmLogin: TMemo;
    Panel3: TPanel;
    edPid: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    edSpcId: TEdit;
    edODT: TEdit;
    edONO: TEdit;
    edECD: TEdit;
    edRES: TEdit;
    btnExec: TButton;
    Label7: TLabel;
    Label8: TLabel;
    edLoc: TEdit;
    Label9: TLabel;
    Label10: TLabel;
    edUser: TEdit;
    lb: TLabel;
    edICD: TEdit;
    edRET: TEdit;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure btnWSelClick(Sender: TObject);
    procedure btnLSelClick(Sender: TObject);
    procedure btnExecClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  F_TEST: TF_TEST;

implementation

{$R *.dfm}

procedure TF_TEST.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action:= caFree;
end;

procedure TF_TEST.FormDestroy(Sender: TObject);
begin
  F_TEST:= nil;
end;

procedure TF_TEST.btnWSelClick(Sender: TObject);
var
  vWork:variant;
  R, i:integer;
begin
  mmWork.Clear;

  R:= ExaminfoList('', edBCD.Text, FormatDateTime('yyyymmdd', now), vWork);
  for i:=0 to R-1 do begin
      mmWork.Lines.Add(vWork[i]);
  end;
end;

procedure TF_TEST.btnLSelClick(Sender: TObject);
var
  vNM:variant;
  R, i:integer;
begin
  mmLogin.Clear;

  R:= UserChk(edID.Text, edPWD.Text, edLoc.Text, vNM);
  for i:=0 to R-1 do begin
      mmLogin.Lines.Add(vNM[i]);
  end;
end;

procedure TF_TEST.btnExecClick(Sender: TObject);
var
  BCD, PID, ODT, ONO, ECD, RES, ICD:string;
  vRES:variant;
  R, i:integer;
begin
  edRET.Text:= ''; 
  BCD:= edSpcId.Text;
  PID:= edPid.Text;
  ODT:= edODT.Text;
  ONO:= edONO.Text;
  ECD:= edECD.Text;
  RES:= edRES.Text;
  ICD:= edICD.Text;

  vRES:= VarArrayCreate([0,0], varVariant);

  vRES[0]:= BCD + '|' + PID + '|' + ODT + '|' + ONO + '|' + ECD + '|' + RES + '|' ;

  R:= ResultList('3', edUser.Text, vRES, ICD, 'N');

  edRET.Text:= IntToStr(R);

end;

end.
