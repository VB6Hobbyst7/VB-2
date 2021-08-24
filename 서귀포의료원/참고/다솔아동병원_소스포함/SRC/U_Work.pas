unit U_Work;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ExtCtrls, Grids, BaseGrid, AdvGrid, Buttons, StdCtrls;

type
  TF_Work = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    StatusBar1: TStatusBar;
    btnClear: TSpeedButton;
    SpeedButton2: TSpeedButton;
    gdWk: TAdvStringGrid;
    btnWK: TSpeedButton;
    Panel5: TPanel;
    dtpWf: TDateTimePicker;
    ckbxAdd: TCheckBox;
    procedure gdWkCanEditCell(Sender: TObject; ARow, ACol: Integer;
      var CanEdit: Boolean);
    procedure gdWkDblClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure btnWKClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnClearClick(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure AddWork(PID, PNM, BCD, ADT, ANO, SPC: string);
    function GetWorkBcd:string;
  end;

const
  W_CKB = 0;
  W_ADT = W_CKB+1;
  W_ANO = W_ADT+1;
  W_BCD = W_ANO+1;
  W_PID = W_BCD+1;
  W_PNM = W_PID+1;
  W_SPC = W_PNM+1;
var
  F_Work: TF_Work;

implementation

uses U_DM, U_IFClass, SetDataBase, GlobalVar, stringLib;

{$R *.dfm}

procedure TF_Work.AddWork(PID, PNM, BCD, ADT, ANO, SPC: string);
var
  ARow:integer;
begin
  if BCD = '' then exit;

  if (gdWk.RowCount = 2) and (gdWk.Cells[W_BCD,1] = '') then
      ARow:= 1
  else begin
      gdWk.AddRow;
      ARow:= gdWk.RowCount -1;
  end;

  gdWk.AddCheckBox(W_CKB, ARow, True, False);

  gdWk.Cells[W_ADT, ARow]:= ADT;
  gdWk.Cells[W_ANO, ARow]:= ANO;
  gdWk.Cells[W_PID, ARow]:= PID;
  gdWk.Cells[W_BCD, ARow]:= BCD;
  gdWk.Cells[W_PNM, ARow]:= PNM;
  gdWk.Cells[W_SPC, ARow]:= SPC;

end;

procedure TF_Work.gdWkCanEditCell(Sender: TObject; ARow, ACol: Integer;
  var CanEdit: Boolean);
begin
  if (ARow > 0) and (ACol = 0) then
      CanEdit:= True
  else
      CanEdit:= False;
end;

procedure TF_Work.gdWkDblClickCell(Sender: TObject; ARow, ACol: Integer);
begin
  if gdWk.Cells[1, ARow] = '' then exit;

  if (ARow > 0) then begin
      if MessageDlg('선택하신 검체정보를 삭제하시겠습니까?', mtConfirmation, mbOKCancel, 1) = mrCancel then exit;

      if gdWk.RowCount = 2 then begin
          gdWk.ClearNormalCells;
      end
      else begin
          gdWk.ClearRows(ARow,1);
          gdWk.RemoveRowsEx(ARow, 1);
      end;
  end;
end;

function TF_Work.GetWorkBcd: string;
var
  B:boolean;
  i:integer;
begin
  Result:= '';

  for i:=1 to gdWk.RowCount -1 do begin
      B:= False;
      gdWk.GetCheckBoxState(0, i, B);
      if B = True then begin
          Result:= gdWk.Cells[W_BCD, i];
          if gdWk.RowCount = 2 then begin
              gdWk.ClearNormalCells;
          end
          else begin
              gdWk.ClearRows(i, 1);
              gdWk.RemoveRowsEx(i, 1);
          end;
          exit;
      end;
  end;
end;

procedure TF_Work.btnWKClick(Sender: TObject);
var
  R, i, OrdCnt:integer;
  WT, WF, sWork, sOrder:string;
  ICD, BCD, ANO, ADT, PID, SPC, PNM:string;
  slWork:TStringList;
begin
  if ckbxAdd.Checked = False then begin
      gdWK.ClearNormalCells;
      gdWK.Row:=1;
      gdWk.RowCount:=2;
  end;

  WF:= FormatDateTime('yyyymmdd', dtpWf.Date);

  slWork:= TStringList.Create;
  try
      sWork:= DM.DownLoadOrder_DJI_WORK(WF);

      if SvrTEST = True then
          TGlobal.DataLog:= sWork;

      //if Length(Trim(sWork)) < 100 then exit;

      //0                         ACAWNIFH    !! 조회완료.                                                 00013181	황영수	LA2	EB	1404020070	20140402	70
      //sOrder:= Copy(sWork,100, Length(sWork)-100);

      slWork.Text:= sWork;

      for i:=0 to slWork.Count -1 do begin
          sOrder:= slWork.Strings[i];
          PID:= Trim(TokenStr(sOrder, TAB, 1));
          PNM:= Trim(TokenStr(sOrder, TAB, 2));
          SPC:= Trim(TokenStr(sOrder, TAB, 3));
          BCD:= Trim(TokenStr(sOrder, TAB, 4));
          ADT:= Trim(TokenStr(sOrder, TAB, 5));
          ANO:= Trim(TokenStr(sOrder, TAB, 6));

          //PID + TAB + PNM + TAB + SPC + TAB + BCD + TAB + ADT + TAB + ANO + TAB

          //ShowMessage(PID+','+ PNM+','+BCD+','+ADT+','+ANO+','+SPC);

          AddWork(PID, PNM, BCD, ADT, ANO, SPC);
      end;


  finally
      slWork.Free;
  end;

end;

procedure TF_Work.FormCreate(Sender: TObject);
begin
  dtpWf.DateTime:= now;
end;

procedure TF_Work.btnClearClick(Sender: TObject);
begin
  gdWk.ClearNormalCells;
  gdWk.Row:=1;
  gdWk.RowCount:=2;
end;

procedure TF_Work.SpeedButton2Click(Sender: TObject);
begin
  Close;
end;

end.
