unit U_ENV;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls;

type
  TF_ENV = class(TForm)
    ColorDialog1: TColorDialog;
    GroupBox1: TGroupBox;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    shMEAN: TShape;
    shRES: TShape;
    sh1SD: TShape;
    sh2SD: TShape;
    sh3SD: TShape;
    GroupBox2: TGroupBox;
    Panel6: TPanel;
    Panel7: TPanel;
    edHospNm: TEdit;
    edCor: TEdit;
    btnSave: TButton;
    Button1: TButton;
    procedure shRESMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure ChangeColor(TSH:TObject; C:TColor);
  end;

var
  F_ENV: TF_ENV;

implementation

uses GlobalVar;

{$R *.dfm}

procedure TF_ENV.ChangeColor(TSH:TObject; C:TColor);
var
  SH:TShape absolute TSH;
  CopNm:string;
begin
  SH.Brush.Color:= C;
  CopNm:= UpperCase(SH.Name);

  if CopNm = 'SHRES' then
      TGlobal.FGrpColor.Res:= C
  else if CopNm = 'SH1SD' then
      TGlobal.FGrpColor.Sd1:= C
  else if CopNm = 'SH2SD' then
      TGlobal.FGrpColor.Sd2:= C
  else if CopNm = 'SH3SD' then
      TGlobal.FGrpColor.Sd3:= C
  else if CopNm = 'SHMEAN' then
      TGlobal.FGrpColor.Mean:= C;
end;

procedure TF_ENV.shRESMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if ColorDialog1.Execute then begin
      ChangeColor(Sender, ColorDialog1.Color);
  end;
end;

procedure TF_ENV.FormCreate(Sender: TObject);
begin
  shRES.Brush.Color:= TGlobal.FGrpColor.Res;
  shMEAN.Brush.Color:= TGlobal.FGrpColor.Mean;
  sh1SD.Brush.Color:= TGlobal.FGrpColor.Sd1;
  sh2SD.Brush.Color:= TGlobal.FGrpColor.Sd2;
  sh3SD.Brush.Color:= TGlobal.FGrpColor.Sd3;
end;

procedure TF_ENV.Button1Click(Sender: TObject);
begin
  Close;
end;

procedure TF_ENV.btnSaveClick(Sender: TObject);
begin
  TGlobal.FHospNm:= edHospNm.Text;
  TGlobal.FCor   := edCor.Text;
  TGlobal.SaveIni;
end;

end.
