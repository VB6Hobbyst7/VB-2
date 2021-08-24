unit U_CodeM;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, Grids, BaseGrid, AdvGrid, ComCtrls, ExtCtrls,
  ADODB, Spin;
type
  TF_CodeM = class(TForm)
    Panel1: TPanel;
    StatusBar1: TStatusBar;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    gdCodeM: TAdvStringGrid;
    btnAdd: TSpeedButton;
    btnDel: TSpeedButton;
    Panel5: TPanel;
    Panel6: TPanel;
    Panel7: TPanel;
    Panel8: TPanel;
    Panel9: TPanel;
    edCode: TEdit;
    edName: TEdit;
    edAbbr: TEdit;
    btnClose: TSpeedButton;
    btnView: TSpeedButton;
    Panel10: TPanel;
    Panel11: TPanel;
    edUpCode: TEdit;
    btnSave: TBitBtn;
    Panel17: TPanel;
    edLow: TEdit;
    edHigh: TEdit;
    seSeq: TSpinEdit;
    Panel12: TPanel;
    edOrdcd: TEdit;
    Panel13: TPanel;
    edPanel: TEdit;
    Panel14: TPanel;
    edIfcd: TEdit;
    procedure NextSelect(Sender:TObject; var Key:Char);
    procedure FormShow(Sender: TObject);
    procedure gdCodeMClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure btnDelClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure btnViewClick(Sender: TObject);
    procedure gdCodeMClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnAddClick(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure edCodeKeyPress(Sender: TObject; var Key: Char);
    procedure edNameKeyPress(Sender: TObject; var Key: Char);
    procedure edUpCodeKeyPress(Sender: TObject; var Key: Char);
    procedure edAbbrKeyPress(Sender: TObject; var Key: Char);
    procedure edLowKeyPress(Sender: TObject; var Key: Char);
    procedure edHighKeyPress(Sender: TObject; var Key: Char);
    procedure edSeqKeyPress(Sender: TObject; var Key: Char);
    procedure edCode1KeyPress(Sender: TObject; var Key: Char);
    procedure codenm1KeyPress(Sender: TObject; var Key: Char);
    procedure edCode2KeyPress(Sender: TObject; var Key: Char);
    procedure codenm2KeyPress(Sender: TObject; var Key: Char);
    procedure edCode3KeyPress(Sender: TObject; var Key: Char);
    procedure codenm3KeyPress(Sender: TObject; var Key: Char);
    procedure edCode4KeyPress(Sender: TObject; var Key: Char);
    procedure codenm4KeyPress(Sender: TObject; var Key: Char);
    procedure edCode5KeyPress(Sender: TObject; var Key: Char);
    procedure codenm5KeyPress(Sender: TObject; var Key: Char);
  private
    procedure VeiwCodePanel(OCode, sPanel, sSeq,sCode,sName,sAbbr,sUp,sIf,sL,sH:string);
    procedure ClearCodePanel;
    procedure ViewCodeList;
  public
  end;

var
  F_CodeM: TF_CodeM;

implementation

uses SetDataBase, StringLib, U_DM;

{$R *.dfm}

const
  C_Seq = 0;
  C_ORD = C_SEQ+1;
  C_PNL = C_ORD+1;
  C_Code = C_PNL+1;
  C_Name = C_Code+1;
  C_Abbr = C_Name+1;
  C_IFCD = C_Abbr+1;
  C_UPCd = C_IFCD+1;
  C_RefL = C_UPCd+1;
  C_RefH = C_RefL+1;


procedure TF_CodeM.FormShow(Sender: TObject);
begin
  gdCodeM.HideColumns(C_Abbr, C_RefH);

  btnView.Click;

  gdCodeM.OnClickCell(Sender, 1,1);
end;

procedure TF_CodeM.ViewCodeList;
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
          AddSql(' Select * From TB_CodeInfo  ');
          AddSQl(' Order By DispSeq');
          RCount:= LocalSelect(QryEx);

          if RCount = 0 then begin
              ShowMessage('설정된 검사코드가 없습니다!');
              exit;
          end;

          gdCodeM.RowCount:= RCount+1;

          i:=0;
          with QryEx do begin
              While Not Eof do begin
                  inc(i);
                  gdCodeM.Cells[C_ORD, i]:= FieldByName('OrdCode').AsString;
                  gdCodeM.Cells[C_PNL, i]:= FieldByName('Panel').AsString;
                  gdCodeM.Cells[C_Seq, i] := IntToStr(FieldByName('DispSeq').AsInteger);
                  gdCodeM.Cells[C_Code, i]:= FieldByName('ExamCode').AsString;
                  gdCodeM.Cells[C_Name, i]:= FieldByName('ExamName').AsString;
                  gdCodeM.Cells[C_Abbr, i]:= FieldByName('Abbr').AsString;
                  gdCodeM.Cells[C_RefL, i]:= FloatToStr(FieldByName('RefLow').AsFloat);
                  gdCodeM.Cells[C_RefH, i]:= FloatToStr(FieldByName('RefHigh').AsFloat);
                  gdCodeM.Cells[C_UPCd, i]:= FieldByName('UpCode').AsString;
                  gdCodeM.Cells[C_IFCd, i]:= FieldByName('IFCode').AsString;

                  Next;
              end;
          end;
      end;

  finally
      QryEx.Free;
      TSql.Free;
  end;
end;

procedure TF_CodeM.gdCodeMClickCell(Sender: TObject; ARow, ACol: Integer);
var
  sCode,sName,sAbbr,OCode,sPanel,
  sSeq,sLow,sHigh,sUp,sIF:string;
begin
  if ARow > 0 then begin
      OCode:= gdCodeM.Cells[C_ORD, ARow];
      sPanel:=gdCodeM.Cells[C_PNL, ARow];
      sCode:= gdCodeM.Cells[C_Code, ARow];
      sName:= gdCodeM.Cells[C_Name, ARow];
      sAbbr:= gdCodeM.Cells[C_Abbr, ARow];
      sUp  := gdCodeM.Cells[C_UPCD, ARow];
      sIF  := gdCodeM.Cells[C_IFCD, ARow];
      sSeq := gdCodeM.Cells[C_SEQ,  ARow];
      sLow := gdCodeM.Cells[C_RefL, ARow];
      sHigh:= gdCodeM.Cells[C_RefH, ARow];

      VeiwCodePanel(OCode, sPanel, sSeq, sCode, sName, sAbbr, sUp, sIf, sLow, sHigh );
  end;
end;

procedure TF_CodeM.VeiwCodePanel(OCode, sPanel, sSeq, sCode, sName, sAbbr, sUp, sIf, sL,
  sH: string);
begin
  ClearCodePanel;

  edOrdcd.Text:= OCode;
  edPanel.Text:= sPanel;
  seSeq.Text:= sSeq;
  edCode.Text:= sCode;
  edName.Text:= sName;
  edAbbr.Text:= sAbbr;
  edUpCode.Text:= sUp;
  edIfcd.Text:= sIf;
  edLow.Text := sL;
  edHigh.Text:= sH;
end;

procedure TF_CodeM.ClearCodePanel;
begin
  edOrdcd.Text:='';
  edPanel.Text:='';
  seSeq.Text := '';
  edCode.Text:= '';
  edName.Text:= '';
  edAbbr.Text:= '';
  edUpCode.Text:='';
  edLow.Text := '';
  edHigh.Text:= '';

  edCode.SetFocus;
end;

procedure TF_CodeM.btnDelClick(Sender: TObject);
var
  sCode,OCode:string;
begin
  sCode:= edCode.Text;
  OCode:= edOrdcd.Text;

  if (Trim(OCode) = '') or (Trim(sCode) = '' ) then begin
      ShowMessage('검사코드를 선택하세요!');
      exit;
  end;

  if MessageDlg('해당 코드를 삭제하시겠습니까? 삭제하시면 해당 코드 결과를 사용할수 없습니다!', mtInformation, mbOKCancel, 1) = mrOk then
  begin
      if DM.DeleteOneCode(OCode, sCode) then begin
          btnView.Click;
      end
      else begin
          ShowMessage('삭제 실패!');
      end;
  end;
end;

procedure TF_CodeM.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action:= caFree;
end;

procedure TF_CodeM.FormDestroy(Sender: TObject);
begin
  F_CodeM:= nil;
end;

procedure TF_CodeM.btnViewClick(Sender: TObject);
begin
  ViewCodeList;
end;

procedure TF_CodeM.gdCodeMClick(Sender: TObject);
begin
  gdCodeM.OnClickCell(Sender, gdCodeM.Row,gdCodeM.Col);
end;

procedure TF_CodeM.btnSaveClick(Sender: TObject);
var
  sCode,sName,sAbbr,OCode,sPanel,
  sSeq,sLow,sHigh,sUp,sIf:string;
begin
  OCode:= edOrdcd.Text;
  sPanel:= edPanel.Text;
  sCode:= edCode.Text;
  sName:= edName.Text;
  sAbbr:= edAbbr.Text;
  sSEq := seSeq.Text;
  sLow := edLow.Text;
  sHigh:= edHigh.Text;
  sUp  := edUpCode.Text;
  sIf  := edIfcd.Text;

  if Trim(OCode) = '' then begin
      ShowMessage('처방코드를 입력하세요!');
      exit;
  end;

  if Trim(sPanel) = '' then begin
      ShowMessage('검사판넬을 입력하세요!');
      exit;
  end;

  if Trim(sCode) = '' then begin
      ShowMessage('검사코드를 입력하세요!');
      exit;
  end;

  if StrToIntDef(sSeq,-1) < 0 then begin
      ShowMessage('순서를 입력하세요!');
      exit;
  end;

  if Trim(sUp) = '' then begin
      ShowMessage('장비 수신코드를 입력하세요!');
      edUpCode.SetFocus;
      exit;
  end;

  if Trim(sIf) = '' then begin
      ShowMessage('장비 전송코드를 입력하세요!');
      edIfcd.SetFocus;
      exit;
  end;

  if Trim(sAbbr) = '' then begin
      ShowMessage('검사약어를 입력하세요!');
      edAbbr.SetFocus;
      exit;
  end;

  if (StrToFloatDef(sLow,-100) < -99) or
     (StrToFloatDef(sHigh,-100) < -99) then begin
      ShowMessage('잘못된 상한치 or 하한치 값입니다.');
      exit;
  end;

  if DM.SaveOneCode(OCode, sPanel, sCode,sName,sAbbr, sIf, sUp,sLow,sHigh,sSeq) then begin
      btnView.Click;
  end
  else
      ShowMessage('저장실패');

end;

procedure TF_CodeM.btnAddClick(Sender: TObject);
begin
  ClearCodePanel;
end;

procedure TF_CodeM.btnCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TF_CodeM.edCodeKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      SelectNext(Sender As TWinControl, True, True);
  end;
end;

procedure TF_CodeM.edNameKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      SelectNext(Sender As TWinControl, True, True);
  end;
end;

procedure TF_CodeM.edUpCodeKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      SelectNext(Sender As TWinControl, True, True);
  end;
end;

procedure TF_CodeM.edAbbrKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      SelectNext(Sender As TWinControl, True, True);
  end;
end;

procedure TF_CodeM.edLowKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      SelectNext(Sender As TWinControl, True, True);
  end;
end;

procedure TF_CodeM.edHighKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      SelectNext(Sender As TWinControl, True, True);
  end;
end;

procedure TF_CodeM.edSeqKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      SelectNext(Sender As TWinControl, True, True);
  end;
end;

procedure TF_CodeM.edCode1KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      SelectNext(Sender As TWinControl, True, True);
  end;
end;

procedure TF_CodeM.codenm1KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      SelectNext(Sender As TWinControl, True, True);
  end;
end;

procedure TF_CodeM.edCode2KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      SelectNext(Sender As TWinControl, True, True);
  end;
end;

procedure TF_CodeM.codenm2KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      SelectNext(Sender As TWinControl, True, True);
  end;
end;

procedure TF_CodeM.edCode3KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      SelectNext(Sender As TWinControl, True, True);
  end;
end;

procedure TF_CodeM.codenm3KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      SelectNext(Sender As TWinControl, True, True);
  end;
end;

procedure TF_CodeM.edCode4KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      SelectNext(Sender As TWinControl, True, True);
  end;
end;

procedure TF_CodeM.codenm4KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      SelectNext(Sender As TWinControl, True, True);
  end;
end;

procedure TF_CodeM.edCode5KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      SelectNext(Sender As TWinControl, True, True);
  end;
end;

procedure TF_CodeM.codenm5KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      SelectNext(Sender As TWinControl, True, True);
  end;
end;

procedure TF_CodeM.NextSelect(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      SelectNext(Sender As TWinControl, True, True);
  end;

end;

end.
