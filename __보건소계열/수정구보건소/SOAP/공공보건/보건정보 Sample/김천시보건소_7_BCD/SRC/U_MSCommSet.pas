unit U_MSCommSet;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls;

type
  TF_MSCommSet = class(TForm)
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
    Button2: TButton;
    Button1: TButton;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  F_MSCommSet: TF_MSCommSet;

implementation

{$R *.dfm}

end.
