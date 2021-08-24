unit U_CODE_SET;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, BaseGrid, AdvGrid, Buttons, ExtCtrls;

type
  TF_CodeSet = class(TForm)
    Panel1: TPanel;
    btnClose: TSpeedButton;
    btnView: TSpeedButton;
    Panel3: TPanel;
    gdCodeM: TAdvStringGrid;
    Panel4: TPanel;
    btnAdd: TSpeedButton;
    btnDel: TSpeedButton;
    btnSave: TSpeedButton;
    Panel5: TPanel;
    Panel6: TPanel;
    Panel7: TPanel;
    Panel8: TPanel;
    Panel9: TPanel;
    edCode: TEdit;
    edName: TEdit;
    edAbbr: TEdit;
    Panel10: TPanel;
    Panel11: TPanel;
    edUpCode: TEdit;
    edSub: TEdit;
    Panel12: TPanel;
    edLow: TEdit;
    edHigh: TEdit;
    edSeq: TEdit;
    procedure gdCodeMClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure FormShow(Sender: TObject);
    procedure btnDelClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure btnViewClick(Sender: TObject);
    procedure gdCodeMClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnAddClick(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure MySelectNext(Sender:TObject; var Key:Char);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure VeiwCodePanel(sSeq,sCode,sName,sAbbr,sUp,sSub,sL,sH:string);
    procedure ClearCodePanel;
    procedure ViewCodeList;
  end;

var
  F_CodeSet: TF_CodeSet;

implementation

uses SetDataBase, StringLib, U_DM, GlobalVar, U_CodeInfo, ADODB;

{$R *.dfm}

const
  C_Seq  = 0;
  C_Code = C_Seq+1;
  C_Name = C_Code+1;
  C_Abbr = C_Name+1;
  C_UPCd = C_Abbr+1;
  C_SUB  = C_UPCD+1;
  C_RefL = C_SUB+1;
  C_RefH = C_RefL+1;



{ TF_CodeSet }

procedure TF_CodeSet.ClearCodePanel;
begin
  edSeq.Text := '';
  edCode.Text:= '';
  edName.Text:= '';
  edAbbr.Text:= '';
  edUpCode.Text:='';
  edLow.Text := '';
  edHigh.Text:= '';
  edSub.Text:='';

end;

procedure TF_CodeSet.VeiwCodePanel(sSeq, sCode, sName, sAbbr, sUp, sSub,
  sL, sH: string);
begin
  ClearCodePanel;
  edSeq.Text:= sSeq;
  edCode.Text:= sCode;
  edName.Text:= sName;
  edAbbr.Text:= sAbbr;
  edUpCode.Text:= sUp;
  edLow.Text := sL;
  edHigh.Text:= sH;
  edSub.Text := sSub;

end;

procedure TF_CodeSet.ViewCodeList;
var
  TSql:TQueryInfo;
  QryEx:TAdoQuery;
  i:integer;
begin
  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Self);
  try
      with TSql do begin
          Clear;
          AddSql(' Select * From TB_Code Order By DispSeq');
          RCount:= LocalSelect(QryEx);

          if RCount = 0 then begin
              ShowMessage('��ϵ� �ڵ尡 �����ϴ�!');
              exit;
          end;

          gdCodeM.RowCount:= RCount+1;

          i:=0;
          with QryEx do begin
              While Not Eof do begin
                  inc(i);
                  gdCodeM.Cells[C_Seq, i] := PadLeftStr(IntToStr(FieldByName('DispSeq').AsInteger), '0', 3);
                  gdCodeM.Cells[C_Code, i]:= FieldByName('ExamCode').AsString;
                  gdCodeM.Cells[C_Name, i]:= FieldByName('ExamName').AsString;
                  gdCodeM.Cells[C_Abbr, i]:= FieldByName('Abbr').AsString;
                  gdCodeM.Cells[C_RefL, i]:= FloatToStr(FieldByName('RefLow').AsFloat);
                  gdCodeM.Cells[C_RefH, i]:= FloatToStr(FieldByName('RefHigh').AsFloat);
                  gdCodeM.Cells[C_UPCd, i]:= FieldByName('IFCode').AsString;
                  gdCodeM.Cells[C_SUB,  i]:= FieldByName('IFCode_Sub').AsString;
                  Next;
              end;
          end;
      end;

  finally
      QryEx.Free;
      TSql.Free;
  end;

end;

procedure TF_CodeSet.gdCodeMClickCell(Sender: TObject; ARow,
  ACol: Integer);
var
  sCode,sName,sAbbr,sLoc,
  sSeq,sLow,sHigh,sUp,
  SubCd:string;
begin
  if ARow > 0 then begin
      sCode:= gdCodeM.Cells[C_Code, ARow];
      sName:= gdCodeM.Cells[C_Name, ARow];
      sAbbr:= gdCodeM.Cells[C_Abbr, ARow];
      sUp  := gdCodeM.Cells[C_UPCD, ARow];
      SubCd:= gdCodeM.Cells[C_SUB , ARow];
      sSeq := gdCodeM.Cells[C_SEQ,  ARow];
      sLow := gdCodeM.Cells[C_RefL, ARow];
      sHigh:= gdCodeM.Cells[C_RefH, ARow];
      VeiwCodePanel(sSeq, sCode, sName, sAbbr, sUp, SubCd, sLow, sHigh);
  end;

end;

procedure TF_CodeSet.FormShow(Sender: TObject);
var
  i:integer;
begin
  gdCodeM.HideColumns(C_Abbr, C_RefH);
  btnView.Click;
  gdCodeM.OnClickCell(Sender, 1,1);

end;

procedure TF_CodeSet.btnDelClick(Sender: TObject);
var
  sCode:string;
begin
  sCode:= edCode.Text;

  if Trim(sCode) = '' then begin
      ShowMessage('�ڵ带 �Է� �ϼž� �մϴ�!');
      exit;
  end;

  if MessageDlg('������ �ڵ带 ���� �Ͻðڽ��ϱ�?', mtInformation, mbOKCancel, 1) = mrOk then
  begin
      if DM.DeleteOneCode(sCode) then begin
          btnView.Click;
      end
      else begin
          ShowMessage('���� ����! Error Log�� Ȯ���ϼ���!');
      end;
  end;

end;

procedure TF_CodeSet.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    Action:= caFree;
end;

procedure TF_CodeSet.FormDestroy(Sender: TObject);
begin
    F_CodeSet:= nil;
end;

procedure TF_CodeSet.btnViewClick(Sender: TObject);
begin
  ViewCodeList;
end;

procedure TF_CodeSet.gdCodeMClick(Sender: TObject);
begin
  gdCodeM.OnClickCell(Sender, gdCodeM.Row,gdCodeM.Col);
end;

procedure TF_CodeSet.btnSaveClick(Sender: TObject);
var
  sCode,sName,sAbbr,SubCd,
  sSeq,sLow,sHigh,sUp:string;
begin
  sCode:= edCode.Text;
  sName:= edName.Text;
  sAbbr:= edAbbr.Text;
  sSEq := edSeq.Text;
  sLow := edLow.Text;
  sHigh:= edHigh.Text;
  sUp  := edUpCode.Text;
  SubCd:= edSub.Text;

  if Trim(sCode) = '' then begin
      ShowMessage('�ڵ带 �Է� �ϼž� �մϴ�!');
      exit;
  end;

  if StrToIntDef(sSeq,-1) < 0 then begin
      ShowMessage('�˻������ �߸��� ���Դϴ�!');
      exit;
  end;

  if Trim(sUp) = '' then begin
      ShowMessage('�����ڵ尡 ������ ����� �����Ҽ� �����ϴ�!');
      edUpCode.SetFocus;
      exit;
  end;

  if Trim(SubCd) = '' then begin
      ShowMessage('SUB �ڵ尡 ������ QC ����� �����Ҽ� �����ϴ�!');
      edSub.SetFocus;
      exit;
  end;

  if (StrToFloatDef(sLow,-100) < -99) or
     (StrToFloatDef(sHigh,-100) < -99) then begin
      ShowMessage('�߸��� ����ġ ���Դϴ�!');
      exit;
  end;

  if Trim(sAbbr) = '' then begin
      if MessageDlg('�˻�� �������� ������ ȭ�鿡 ǥ�õ��� �ʽ��ϴ�!'+#13#10+'��� �Ͻðڽ��ϱ�?',
                     mtWarning, mbOKCancel, 1) = mrOk then begin
          if DM.SaveOneCode(sCode,sName,sAbbr,sUp,SubCd,sLow,sHigh,sSeq) then begin
              btnView.Click;
          end
          else begin
              ShowMessage('�������! Error Log�� Ȯ���ϼ���!');
          end;
      end
      else begin
          exit;
      end;
  end
  else begin
      if DM.SaveOneCode(sCode,sName,sAbbr,sUp,SubCd,sLow,sHigh,sSeq) then begin
          btnView.Click;
      end
      else
          ShowMessage('�������! Error Log�� Ȯ���ϼ���!');
  end;

end;

procedure TF_CodeSet.btnAddClick(Sender: TObject);
begin
  ClearCodePanel;
end;

procedure TF_CodeSet.btnCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TF_CodeSet.MySelectNext(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:=#0;
      SelectNext(Sender AS TWincontrol, True, True);
  end;

end;

end.
