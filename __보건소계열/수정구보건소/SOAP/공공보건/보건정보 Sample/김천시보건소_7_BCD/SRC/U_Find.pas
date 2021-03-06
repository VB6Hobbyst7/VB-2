unit U_Find;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, Buttons, Grids, BaseGrid, AdvGrid, ExtCtrls,
  ADODB, U_Main;

type
  TExamView = class(TObject)
    FromDt,
    ToDt:TDateTime;
    iOptID,
    iOptRslt,
    iOptSend,
    iOptLoc:integer;
    sOptId,
    sOtpRslt,
    sOptSend,
    sOptLoc:string;
    sBarcode,
    sPatNo:string;
    bBcd,
    bPat:boolean;
    constructor Create;
  end;

type
  TF_Find = class(TForm)
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    gdMaster: TAdvStringGrid;
    gdResult: TAdvStringGrid;
    Panel1: TPanel;
    GroupBox1: TGroupBox;
    Panel10: TPanel;
    Panel11: TPanel;
    edBarCode: TEdit;
    edPatNo: TEdit;
    GroupBox2: TGroupBox;
    Panel6: TPanel;
    dtpFrom: TDateTimePicker;
    Label1: TLabel;
    dtpTo: TDateTimePicker;
    Panel7: TPanel;
    cmbxID: TComboBox;
    Panel8: TPanel;
    cmbxSend: TComboBox;
    Panel9: TPanel;
    cmbxResult: TComboBox;
    btnView: TSpeedButton;
    btnFind: TSpeedButton;
    ckbxAll: TCheckBox;
    Panel12: TPanel;
    cmbxLoc: TComboBox;
    Label2: TLabel;
    lbMsg: TLabel;
    rbtBcd: TRadioButton;
    rbtPat: TRadioButton;
    btnDel: TSpeedButton;
    btnClose: TSpeedButton;
    pnData: TPanel;
    Panel13: TPanel;
    Panel14: TPanel;
    Panel15: TPanel;
    Panel16: TPanel;
    Panel17: TPanel;
    Panel18: TPanel;
    Panel19: TPanel;
    edSendSpcid: TEdit;
    edSendPatId: TEdit;
    edLocaion: TEdit;
    edFlag: TEdit;
    edUId: TEdit;
    pnDatetime: TPanel;
    pnSeq: TPanel;
    ProgressBar1: TProgressBar;
    pnCheck: TPanel;
    pnOrdCreate: TPanel;
    pnExcept: TPanel;
    pnUpload: TPanel;
    pnOrdCode: TPanel;
    btnSave: TBitBtn;
    btnSend: TBitBtn;
    btnClosePanel: TBitBtn;
    Panel20: TPanel;
    edLotNo: TEdit;
    pnICode: TPanel;
    procedure FormShow(Sender: TObject);
    procedure ckbxAllClick(Sender: TObject);
    procedure btnFindClick(Sender: TObject);
    procedure btnViewClick(Sender: TObject);
    procedure gdMasterCanEditCell(Sender: TObject; ARow, ACol: Integer;
      var CanEdit: Boolean);
    procedure gdMasterClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure gdMasterClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure gdMasterGetAlignment(Sender: TObject; ARow, ACol: Integer;
      var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure gdResultGetAlignment(Sender: TObject; ARow, ACol: Integer;
      var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure gdResultGetCellColor(Sender: TObject; ARow, ACol: Integer;
      AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
    procedure gdMasterGetCellColor(Sender: TObject; ARow, ACol: Integer;
      AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
    procedure btnDelClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure btnSendClick(Sender: TObject);
    procedure btnOrderClick(Sender: TObject);
    procedure gdMasterDblClick(Sender: TObject);
    procedure btnClosePanelClick(Sender: TObject);
    procedure edFlagChange(Sender: TObject);
    procedure edSendSpcidKeyPress(Sender: TObject; var Key: Char);
    procedure edSendPatIdKeyPress(Sender: TObject; var Key: Char);
    procedure edLocaionKeyPress(Sender: TObject; var Key: Char);
    procedure edFlagKeyPress(Sender: TObject; var Key: Char);
    procedure edUIdKeyPress(Sender: TObject; var Key: Char);
    procedure edSendPatIdKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure edLocaionKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure edFlagKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure edUIdKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure edSendSpcidKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure pnDataMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure btnSend_GridClick(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure gdMasterGridHint(Sender: TObject; ARow, ACol: Integer;
      var hintstr: String);
    procedure edLotNoKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure edLotNoKeyPress(Sender: TObject; var Key: Char);
  private
    procedure ClearGrid(var TGrid:TAdvStringGrid);
    procedure ViewExamList(TView:TExamView);
    procedure FindSpc(TView:TExamView);
    procedure ViewResult(ExamDate:string; ExamSeq:integer; QCYN:boolean);
    procedure SetDefaultPanelColor(clDef:TColor);
  public

  end;

var
  F_Find: TF_Find;

implementation

uses SetDataBase, U_DM, U_CodeInfo, GlobalVar, U_IFClass, DB;

const
  C_Check = 0;
  C_Date  = C_Check+1;
  C_Seq   = C_Date+1;
  C_SpcID = C_Seq+1;
  C_PatID = C_SpcID+1;
  C_Flag  = c_PatId+1;
  C_IdStat= C_Flag+1;
  C_State = C_IdStat+1;
  C_LotNo = C_State+1;
  C_Locat = C_LotNo+1;
  C_User  = C_Locat+1;
  C_ICode = C_User+1;

  R_Code = 0;
  R_Name = R_Code+1;
  R_Rslt = R_Name+1;
  R_Low  = R_Rslt+1;
  R_High = R_Low+1;
  R_Delta = R_High+1;
  R_Panic = R_Delta+1;
  R_Criti = R_Panic+1;
  R_LowHigh = R_Criti+1;
  R_UPCD    = R_LowHigh+1;

{$R *.dfm}

{ TExamView }

constructor TExamView.Create;
begin
  with F_Find do begin
      FromDt:= dtpFrom.DateTime;
      ToDt  := dtpTo.DateTime;

      iOptID := cmbxID.ItemIndex;
      iOptRslt:= cmbxResult.ItemIndex;
      iOptSend:= cmbxSend.ItemIndex;
      iOptLoc := cmbxLoc.ItemIndex;

      sOptId:= cmbxId.Text;
      sOtpRslt:= cmbxResult.Text;
      sOptSend:= Copy(cmbxSend.Text,1,1);
      sOptLoc:= cmbxLoc.Text;
      
      sBarcode:= edBarCode.Text;
      sPatNo  := edPatNo.Text;
      bBcd:= rbtBcd.Checked;
      bPat:= rbtPat.Checked;
  end;
end;

{ TF_Find }

procedure TF_Find.ClearGrid(var TGrid: TAdvStringGrid);
begin
  with TGrid do begin
      TGrid.ClearNormalCells;
      TGrid.RowCount:=2;
  end;
end;

procedure TF_Find.FindSpc(TView: TExamView);
var
  TSql:TQueryInfo;
  QryEx:TAdoQuery;
  i:integer;
  sDateTime:string;
  dDate:TDateTime;
begin
  ClearGrid(gdMaster);
  ClearGrid(gdResult);

  TSql:= TQueryInfo.Create;
  QryEx:= TADOQuery.Create(Self);

  try
      with TSql, TView do begin

          Clear;
          AddSql(' Select * From TB_Master ');

          if TView.bBcd then
              AddSql(' Where BarCode = '''+sBarcode+''' ')
          else
              AddSql(' Where PatNo = '''+sPatNo+''' ');

          RCount:= LocalSelect(QryEx);

          if RCount = 0 then
              exit;

          gdMaster.RowCount:= RCount +1;

          i:=0;
          with QryEx do begin
              while Not Eof do begin
                  inc(i);
                  gdMaster.Cells[C_Seq  , i] := IntToStr(FieldByName('ExamSeq').AsInteger);
                  gdMaster.AddCheckBox(C_Check,i,False,False);
                  sDateTime:= FieldByName('ExamDate').AsString + FieldByName('ExamTime').AsString;
                  dDate:= GetSampleDateTime(sDateTime);
                  gdMaster.Cells[C_Date , i] := FormatDateTime('yyyy-mm-dd hh:nn:ss', dDate);
                  gdMaster.Cells[C_Spcid , i]:= FieldByName('BarCode').AsString;
                  gdMaster.Cells[C_PatID , i]:= FieldByName('PatNo').AsString;
                  gdMaster.Cells[C_Flag , i] := FieldByName('Flag').AsString;
                  gdMaster.Cells[C_Locat , i]:= FieldByName('Location').AsString;
                  gdMaster.Cells[C_IdStat, i]:= FieldByName('IDState').AsString;
                  gdMaster.Cells[C_State , i]:= FieldByName('ErrState').AsString;
                  gdMaster.Cells[C_LotNo , i]:= FieldByName('LotNo').AsString;
                  gdMaster.Cells[C_User  , i]:= FieldByName('UserID').AsString;
                  gdMaster.Cells[C_ICode, i]:= FieldByName('ICODE').AsString;
                  Next;
              end;
          end;

      end;

  finally
      QryEx.Free;
      TSql.Free;
  end;

end;

procedure TF_Find.ViewExamList(TView: TExamView);
var
  TSql:TQueryInfo;
  QryEx:TAdoQuery;
  cFrom, cTo, sDateTime:string;
  i:integer;
  dDate:TDateTime;
begin
  ClearGrid(gdMaster);
  ClearGrid(gdResult);

  cFrom:= FormatDateTime('yyyymmdd', TView.FromDt);
  cTo  := FormatDateTime('yyyymmdd', TView.ToDt);

  TSql:= TQueryInfo.Create;
  QryEx:= TADOQuery.Create(Self);

  try
      with TSql, TView do begin

          Clear;
          AddSql(' Select * From TB_Master ');
          AddSql(' Where ExamDate Between '''+cFrom+''' And '''+cTo+''' ');

          if iOptId <> 0 then
              AddSql(' And IdState = '''+Copy(sOptId,1,1)+''' ');

          Case iOptRslt of
              1: AddSql(' And (Delta = ''Y'' or Panic = ''Y'' or Critical = ''Y'' ) ');
              2: AddSql(' And (Delta is null and Panic is null and Critical is null ) ');
          end;

          Case iOptSend of
              0:;
              1: AddSql(' And ErrState <> ''C'' ');
              else
                  AddSql(' And ErrState = '''+sOptSend+''' ');
          end;

          if iOptLoc <> 0 then
              AddSql(' And Location = '''+sOptLoc+''' ');

          AddSql(' Order By ExamDate, ExamSeq ');

          RCount:= LocalSelect(QryEx);

          if RCount = 0 then
              exit;

          gdMaster.RowCount:= RCount +1;

          i:=0;
          with QryEx do begin
              while Not Eof do begin
                  inc(i);
                  gdMaster.Cells[C_Seq  , i] := IntToStr(FieldByName('ExamSeq').AsInteger);
                  gdMaster.AddCheckBox(C_Check,i,False,False);
                  sDateTime:= FieldByName('ExamDate').AsString + FieldByName('ExamTime').AsString;
                  dDate:= GetSampleDateTime(sDateTime);
                  gdMaster.Cells[C_Date , i] := FormatDateTime('yyyy-mm-dd hh:nn:ss', dDate);
                  gdMaster.Cells[C_Spcid , i]:= FieldByName('BarCode').AsString;
                  gdMaster.Cells[C_PatID , i]:= FieldByName('PatNo').AsString;
                  gdMaster.Cells[C_Flag , i] := FieldByName('Flag').AsString;
                  gdMaster.Cells[C_Locat , i]:= FieldByName('Location').AsString;
                  gdMaster.Cells[C_IdStat, i]:= FieldByName('IDState').AsString;
                  gdMaster.Cells[C_State , i]:= FieldByName('ErrState').AsString;
                  gdMaster.Cells[C_LotNo , i]:= FieldByName('LotNo').AsString;
                  gdMaster.Cells[C_User  , i]:= FieldByName('UserID').AsString;
                  gdMaster.Cells[C_ICode, i]:= FieldByName('ICODE').AsString;
                  Next;
              end;
          end;

      end;

  finally
      QryEx.Free;
      TSql.Free;
  end;
end;

procedure TF_Find.FormShow(Sender: TObject);
var
  i:integer;
begin
  edBarCode.MaxLength:= BarCodeLen;
  edPatNo.MaxLength:= PatNoLen;
  
  cmbxLoc.Clear;
  cmbxLoc.Items.Add('????');

  for i:=0 to TPCode.ItemCount -1 do begin
      if cmbxLoc.Items.IndexOf(TPCode.Location[i]) < 0 then
          cmbxLoc.Items.Add(TPCode.Location[i]);
  end;

  cmbxID.ItemIndex:=0;
  cmbxSend.ItemIndex:=0;
  cmbxResult.ItemIndex:=0;
  cmbxLoc.ItemIndex:=0;
  dtpFrom.Date:= now;
  dtpTo.Date:= now;
  rbtBcd.Checked:= True;
  
  lbMsg.Caption:='';

  Self.Top:= F_Main.Top+10;
  Self.Left:= F_Main.Left+10;

  gdResult.HideColumn(R_LowHigh);
  gdResult.HideColumn(R_UPCD);
  gdMaster.HideColumn(C_ICODE);
end;

procedure TF_Find.ckbxAllClick(Sender: TObject);
var
  i:integer;
begin
  for i:=1 to gdMaster.RowCount -1 do begin
      gdMaster.SetCheckBoxState(C_Check, i, ckbxAll.Checked);
  end;

end;

procedure TF_Find.btnFindClick(Sender: TObject);
var
  TView:TExamView;
begin
  TView:= TExamView.Create;
  try
      FindSpc(TView);
  finally
      TView.Free;
  end;
end;

procedure TF_Find.btnViewClick(Sender: TObject);
var
  TView:TExamView;
begin
  TView:= TExamView.Create;
  try
      ViewExamList(TView);
  finally
      TView.Free;
  end;
end;

procedure TF_Find.gdMasterCanEditCell(Sender: TObject; ARow, ACol: Integer;
  var CanEdit: Boolean);
begin
  if ARow > 0 then begin
      if ACol in [C_Check] then
          CanEdit:= True
      else
          CanEdit:= False;
  end;
end;

procedure TF_Find.ViewResult(ExamDate: string; ExamSeq: integer; QCYN:boolean);
var
  TSql:TQueryInfo;
  QryEx: TAdoQuery;
  i:integer;
  dLow,dHigh:double;
  sLot,sCode:string;
  sResult:string;
begin
  ClearGrid(gdResult);

  TSql:= TqueryInfo.Create;
  QryEx:= TADOQuery.Create(Self);

  try
      with TSql do begin
          Clear;
          if QCYN = True then begin
              AddSql(' Select M.ErrorTxt, M.LotNo, M.QCYN, R.UpCode, R.ExamCode, R.RsltTxt, R.Delta, R.Panic, R.Critical, ');
              AddSql('         R.LH, E.DispSeq, E.RefLow, E.RefHigh, E.ExamName                                           ');
              AddSql(' From TB_Master AS M,                                                                               ');
              AddSql('      TB_Result AS R,                                                                               ');
              AddSql('      TB_QC     AS E                                                                                ');
              AddSql(' Where R.UpCode   = E.UpCode                                                                        ');
              AddSql('   and M.LotNo    = E.LotNo                                                                         ');
              AddSql('   and M.ExamDate = R.ExamDate                                                                      ');
              AddSql('   And M.ExamSeq  = R.ExamSeq                                                                       ');
              AddSql('   and R.ExamDate = '''+ExamDate+'''                                                                ');
              AddSql('   And R.ExamSeq  = '+IntToStr(ExamSeq)+'                                                           ');
              AddSql('  Order By E.DispSeq                                                                                ');
          end
          else begin
              AddSql(' Select M.ErrorTxt, M.LotNo, M.QCYN, R.UpCode, R.ExamCode, R.RsltTxt, R.Delta, R.Panic, R.Critical, ');
              AddSql('        R.LH, E.DispSeq, E.RefLow, E.RefHigh, E.ExamName  ');
              AddSql(' From TB_Master AS M INNER JOIN   ');
              AddSql('     (TB_Result AS R INNER JOIN   ');
              AddSql('      TB_Code   AS E          ');
              AddSql('      ON R.UpCode   = E.UpCode)   ');
              AddSql('      ON M.ExamDate = R.ExamDate  ');
              AddSql('     And M.ExamSeq  = R.ExamSeq   ');
              AddSql(' Where R.ExamDate = '''+ExamDate+''' ');
              AddSql('   And R.ExamSeq  = '+IntToStr(ExamSeq)+' ');
              AddSql(' Order By E.DispSeq ');
          end;
          RCount:= LocalSelect(QryEx);
          if RCount = 0 then
              exit;

          lbMsg.Caption:= QryEx.FieldByName('ErrorTxt').AsString;

          gdResult.RowCount:= RCount+1;
          i:=0;
          with QryEx do begin
              while Not Eof do begin
                  Inc(i);
                  sCode:= FieldByName('ExamCode').AsString;
                  sLot := FieldByName('LotNo').AsString;

                  gdResult.Cells[R_Code, i]:= sCode;
                  gdResult.Cells[R_Name, i]:= FieldByName('ExamName').AsString;
                  gdResult.Cells[R_Rslt, i]:= FieldByName('RsltTxt').AsString;

                  if FieldByName('QCYN').AsBoolean = True then begin
                      if TQCCode.GetRangeData(dLow,dHigh, sLot, sCode) then begin
                          gdResult.Cells[R_Low , i]:= FloatToStr(dLow);
                          gdResult.Cells[R_High, i]:= FloatToStr(dHigh);
                      end;
                  end
                  else begin
                      gdResult.Cells[R_Low , i]:= FloatToStr(FieldByName('RefLow').AsFloat);
                      gdResult.Cells[R_High, i]:= FloatToStr(FieldByName('RefHigh').AsFloat);
                  end;
                  gdResult.Cells[R_Delta, i]:= FieldByName('Delta').AsString;
                  gdResult.Cells[R_Panic, i]:= FieldByName('Panic').AsString;
                  gdResult.Cells[R_Criti, i]:= FieldByName('Critical').AsString;
                  gdResult.Cells[R_LowHigh, i]:= FieldByName('LH').AsString;
                  gdResult.Cells[R_UPCD, i]:= Bool2Str(FieldByName('UpCode').AsBoolean);
                  Next;
              end;
          end;
      end;

  finally

  end;

end;

procedure TF_Find.gdMasterClickCell(Sender: TObject; ARow, ACol: Integer);
var
  ExamDate:string;
  ExamSeq:integer;
  BarCode, PatNo:string;
  QCYN:boolean;
begin
  ExamDate:= GetGridDate(Trim(gdMaster.Cells[C_Date, ARow]));
  ExamSeq := StrToIntDef(Trim(gdMaster.Cells[C_Seq, ARow]),0);

  edBarCode.Text    := Trim(gdMaster.Cells[C_Spcid, ARow]);
  edPatNo.Text      := Trim(gdMaster.Cells[C_Patid, ARow]);
  pnDatetime.Caption:= Trim(gdMaster.Cells[C_Date, ARow]);
  pnSeq.Caption     := Trim(gdMaster.Cells[C_Seq , ARow]);
  edSendSpcid.Text  := edBarCode.Text;
  edSendPatId.Text  := edPatNo.Text;
  edLocaion.Text    := Trim(gdMaster.Cells[C_Locat, ARow]);
  edFlag.Text       := Trim(gdMaster.Cells[C_Flag, ARow]);
  edUId.Text        := Trim(gdMaster.Cells[C_User, ARow]);
  edLotNo.Text      := Trim(gdMaster.Cells[C_LotNo, ARow]);
  pnICode.Caption   := Trim(gdMaster.Cells[C_ICode, ARow]); 
  pnOrdCode.Caption := TPCode.GetPanelCode(edLocaion.Text, edFlag.Text);

  if gdMaster.Cells[C_IDStat, ARow] = 'Q' then
      QCYN:= True
  else
      QCYN:= False;

  SetDefaultPanelColor(clBtnFace);
  ViewResult(ExamDate, ExamSeq, QCYN);

end;

procedure TF_Find.gdMasterClick(Sender: TObject);
begin
  gdMaster.OnClickCell(Sender, gdMaster.Row, gdMaster.Col);
end;

procedure TF_Find.btnSaveClick(Sender: TObject);
var
  spcid,PatNo:string;
  ExamDate:string;
  ExamSeq:integer;
  TPatData:TOnePatient;
  nRow:integer;
begin
  nRow:= gdMaster.Row;
  if nRow = 0 then begin
      ShowMessage('???? ???????? ??????????!');
      exit;
  end;

  if MessageDlg('???? ???????? ?????????????????', mtConfirmation, mbOKCancel, 1) <> mrOk then
      exit;

  TPatData:= TOnePatient.Create;


  try
      with TPatData do begin
          Spcid    := Trim(edSendSpcid.Text);
          PatId    := Trim(edSendPatId.Text);
          ExamDate := GetGridDate(Trim(pnDateTime.Caption));
          ExamSeq  := StrToIntDef(Trim(pnSeq.Caption),0);
          Location := Trim(edLocaion.Text);
          Flag     := Trim(edFlag.Text);
          UserID   := Trim(edUId.Text);
          if DM.SaveOnePatient(TPatData) then begin
              ShowMessage('????????!');
              gdMaster.Cells[C_Spcid, nRow]:= Spcid;
              gdMaster.Cells[C_PatId, nRow]:= PatId;
              gdMaster.Cells[C_Locat, nRow]:= Location;
              gdMaster.Cells[C_User, nRow]:= UserID;
              gdMaster.Cells[C_Flag, nRow]:= Flag;
              SelectNext(Sender As TWinControl, True, True);
          end;
      end;

  finally
      TPatData.Free;
  end;

end;

procedure TF_Find.gdMasterGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
  HAlign:= taCenter;
  VAlign:= vtaCenter;
end;

procedure TF_Find.gdResultGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
  VAlign:= vtaCenter;

  if ARow > 0 then begin
      if ACol in [R_Code,R_Name] then begin
          HAlign:= taLeftJustify;
      end
      else
          HAlign:= taCenter;
  end
  else begin
      HAlign:= taCenter;
  end;

end;

procedure TF_Find.gdResultGetCellColor(Sender: TObject; ARow,
  ACol: Integer; AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
var
  LH,sLotNo:string;
  QCYN:boolean;
begin
  if ARow > 0 then begin
      if ACol in [R_Delta..R_Criti] then begin
          if gdResult.Cells[ACol, ARow] = 'Y' then begin
              ABrush.Color:= clRed;
              AFont.Style:=[fsBold];
          end;
      end;

      if ACol in [R_Rslt] then begin
          if ACol = R_Rslt then
              AFont.Style:= [fsBold];
          LH:= gdResult.Cells[R_LowHigh, ARow];

          if LH = 'L' then
              AFont.Color:= Color_Low
          else
          if LH = 'H' then
              AFont.Color:= Color_High;
      end;

      if ACol = R_Low then begin
          //ABrush.Color:= Color_Low;
          AFont.Style:= [fsBold];
      end;

      if ACol = R_High then begin
          //ABrush.Color:= Color_High;
          AFont.Style:= [fsBold];
      end;
  end
  else begin
      if ACol in [R_Low, R_High] then begin
          AFont.Size:=8;
      end;
  end;

end;

procedure TF_Find.gdMasterGetCellColor(Sender: TObject; ARow,
  ACol: Integer; AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
begin
  if ARow > 0 then begin
      if ACol = C_IdStat then begin
          AFont.Style:= [fsBold];
          //ABrush.Color:= $00A3C2FE;
      end;
      if ACol = C_State then begin
          AFont.Style:= [fsBold];
          if gdMaster.Cells[C_State,ARow] <> 'C' then
              AFont.Color:= clRed;
      end;
      if ACol in [C_Spcid, C_PatId] then begin
          AFont.Style:= [fsBold];
      end;
  end;
end;

procedure TF_Find.btnDelClick(Sender: TObject);
var
  ExamDate:string;
  ExamSeq:integer;
  i,iCount,iIndex:integer;
  bCheck:boolean;
begin
  iCount:=0;
  for i:=1 to gdMaster.RowCount -1 do begin
      bCheck:= False;
      gdMaster.GetCheckBoxState(C_Check, i, bCheck);
      if bCheck = True then begin
          iCount:= 1;
          Break;
      end;
  end;

  iIndex:=1;
  
  if iCount > 0 then begin
      if MessageDlg('?????? ???????? ????????! '+#13#10+ ' ???????? ???? ?????? ???? ????????????? ',
         mtWarning, mbOKCancel,1 ) = mrOk then
      begin
          for i:= gdMaster.RowCount-1 downto 1 do begin
              bCheck:= False;
              gdMaster.GetCheckBoxState(C_Check, i, bCheck);
              if bCheck = True then begin
                  iIndex:=i;
                  ExamDate:= GetGridDate(Trim(gdMaster.Cells[C_Date, i]));
                  ExamSeq := StrToIntDef(Trim(gdMaster.Cells[C_Seq, i]),0);

                  DM.DeleteOneData(ExamDate, ExamSeq);
                  if (gdMaster.RowCount = 2) then
                      gdMaster.ClearNormalCells
                  else begin
                      gdMaster.RemoveRows(i, 1);
                  end;
              end;
          end;
          if iIndex = 1 then
              gdMaster.OnClickCell(nil, 1, C_Spcid)
          else
              gdMaster.OnClickCell(nil, iIndex-1, C_Spcid);
      end;
  end
  else begin
      ShowMessage('???????? ?????? ???? ?????? ??????!');
  end;

end;

procedure TF_Find.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action:= caFree;
end;

procedure TF_Find.FormDestroy(Sender: TObject);
begin
  F_Find:= nil;
end;

procedure TF_Find.btnSendClick(Sender: TObject);
var
  SId,Pid,Loc,sDate,UID,sFlag,sLot,sICode:string;
  nSeq,nRow:integer;
  TMaster:TIfMaster;
  UpSucc:boolean;
  color_Load, color_Def:TColor;
  sDelta,sPanic,sCriti:string;
begin
  Color_Load:= clLime;
  Color_Def := pnCheck.Color;

  TMaster:= TIfMaster.Create;

  try
      with TMaster do begin
          SId:= Trim(edSendSpcid.Text);
          PId:= Trim(edSendPatId.Text);
          sDate:= GetGridDate(pnDatetime.Caption);
          nSeq:= StrToIntDef(pnSeq.Caption,0);
          Loc := Trim(edLocaion.Text);
          UID := Trim(edUId.Text);
          sFlag:= Trim(edFlag.Text);
          sLot:= Trim(edLotNo.Text);
          sICode:= Trim(pnICode.Caption);

          TMaster.ExamDate:= sDate;
          TMaster.ExamSeq := nSeq;

          if DM.GetLocalResult(TMaster) = False then begin
              ShowMessage('???? ???????? ?????? ????????!');
              exit;
          end;

          //Application.ProcessMessages;
          pnCheck.Color:= Color_Load;
          Application.ProcessMessages;

          if TPCode.GetPOCTYN(Loc) = True then
          //if ( Pos( 'OP',Loc ) > 0 ) or ( Pos( 'ER',Loc ) > 0 ) or ( Pos( 'NICU',Loc ) > 0 ) then
              TMaster.CheckIdStatus(PId, Loc, UID, sLot, sFlag, sICode,'')
          else
              TMaster.CheckIdStatus(SId, Loc, UID, sLot, sFlag, sICode,'');

          UpSucc:= False;

          Case TMaster.Status of
              stBarCode: begin
                             //?????????? ????????.
                             if not GetPatIdorState then begin
                                 ShowMessage(ErrorMessage);
                                 exit;
                             end;

                             pnExcept.Color:= color_Load;
                             Application.ProcessMessages;

                             //????
                             if Not ExcuteRecept then begin
                                 ShowMessage(ErrorMessage);
                                 if tGlobal.DebugMode = Debug then
                                     ShowMessage('???????? ????!');
                                 exit;
                             end;

                             GetDeltaPanic;


                             //??????
                             pnUpload.Color:= color_Load;
                             Application.ProcessMessages;

                             //?????????? ???????? ??????????
                             //???????? ??????????.. 2008.8.18
                             //if Pos( 'ER',Loc ) > 0 then
                             if TMaster.POCTYN = True then
                                 UpSucc:= UploadHost_Poct
                             else
                                 UpSucc:= UploadHost_Normal;

                         end;
              stYesQC  : begin
                             pnUpload.Color:= color_Load;
                             Application.ProcessMessages;

                             UpSucc:=UploadHost_QC;
                         end;
              stYesPoct: begin
                             if Not GetDeptorDrorSect then begin
                                 ShowMessage(ErrorMessage);
                                 exit;
                             end;

                             Application.ProcessMessages;
                             pnOrdCreate.Color:= Color_Load;
                             Application.ProcessMessages;

                             if Not CreateOrder then begin
                                 ShowMessage(ErrorMessage);
                                 exit;
                             end;

                             Application.ProcessMessages;
                             pnExcept.Color:= Color_Load;
                             Application.ProcessMessages;

                             if Not ExcuteSpecnum then begin
                                 ShowMessage(ErrorMessage);
                                 if tGlobal.DebugMode = Debug then
                                     ShowMessage('???????? ????!');
                                 exit;
                             end;

                             if Not ExcuteRecept then begin
                                 ShowMessage(ErrorMessage);
                                 if tGlobal.DebugMode = Debug then
                                     ShowMessage('???? ???? ????!');
                                 exit;
                             end;

                             GetDeltaPanic;

                             UpSucc:= UploadHost_Poct;
                         end;
          end;

          if UpSucc = True then begin
              nRow:= gdMaster.Row;
              DM.SaveOrder(TMaster);
              DM.SaveDeltaPanic(TMaster);
              gdMaster.Cells[C_Spcid, nRow]:= TMaster.BarCode;
              gdMaster.Cells[C_PatId, nRow]:= TMaster.PatId;
              gdMaster.Cells[C_Locat, nRow]:= TMaster.Location;
              gdMaster.Cells[C_User, nRow]:= TMaster.UserID;
              gdMaster.Cells[C_Flag, nRow]:= TMaster.Flag;
              gdMaster.Cells[C_State, nRow]:= 'C';

              Application.ProcessMessages;

              pnUpload.Color:= Color_Load;
              //???????? state?? ??????????.. 2008.08.18
              F_Main.DispOneState(TMaster.ExamDate, 'C', '????????', TMaster.ExamSeq, True);

              Application.ProcessMessages;
              ShowMessage('????????!');

              //???????? ?????? ??????.
              if DpYN = True then
                  ShowMessage('Delta Panic Data!');

           end
           else begin
               pnUpload.Color:= clRed;
               ShowMessage(ErrorMessage);
           end;
      end;

  finally
      TMaster.Free;
  end;
{
  iCount:=0;
  for i:=1 to gdMaster.RowCount -1 do begin
      bCheck:= False;
      gdMaster.GetCheckBoxState(C_Check, i, bCheck);
      if bCheck = True then begin
          iCount:= 1;
          Break;
      end;
  end;

  nCount:=0;

  if iCount > 0 then begin
      if MessageDlg('???? ?????? ?????? ????????????? ',
         mtWarning, mbOKCancel,1 ) = mrOk then
      begin
          for i:= gdMaster.RowCount-1 downto 1 do begin
              bCheck:= False;
              gdMaster.GetCheckBoxState(C_Check, i, bCheck);
              if bCheck = True then begin
                  ExamDate:= GetGridDate(Trim(gdMaster.Cells[C_Date, i]));
                  ExamSeq := StrToIntDef(Trim(gdMaster.Cells[C_Seq, i]),0);
                  sBarCode:= Trim(gdMaster.Cells[C_Spcid, i]);
                  sOpId   := Trim(gdMaster.Cells[C_User, i]);
                  sLoc    := Trim(gdMaster.Cells[C_Locat, i]);
                  sFlag   := Trim(gdMaster.Cells[C_Flag, i]);

                  if gdMaster.Cells[C_IDStat, i] = 'Q' then
                      QCYN:= True
                  else
                      QCYN:= False;

                  //?????? ??????????.
                  TGlobal.IDState:= 'B';

                  if DM.ReSend(ExamDate, sBarCode, sOpId, sLoc, sFlag, ExamSeq, QCYN) then begin
                      gdMaster.Cells[C_State,i]:= 'C';

                      if TGlobal.PatNo <> '' then begin
                          DM.SavePatId(TGlobal.PatNo, ExamDate, ExamSeq);
                          gdMaster.Cells[C_PatID, i]:= TGlobal.PatNo;
                      end;
                      inc(nCount);
                  end;
              end;
          end;

          ShowMessage(IntToStr(nCount)+'?? ????!');
      end;
  end
  else begin
      ShowMessage('?????? ?????? ???? ?????? ??????!');
  end;
  }

  SelectNext(Sender As TWinControl, True, True);
end;

procedure TF_Find.btnOrderClick(Sender: TObject);
var
  ExamDate,PatNo:string;
  ExamSeq:integer;
  Spcid,sOpId,sLoc,sFlag:string;
begin
  PatNo:= Trim(edPatNo.Text);
  if Length(PatNo) <> PatNoLen then begin
      ShowMessage('?????????? ??????????! ???????? ???? ????????.');
      exit;
  end;

  //???????? ????.
  if gdMaster.Cells[C_PatID, gdMaster.Row] <> PatNo then
      btnSave.Click;

  ExamDate:= GetGridDate(pnDateTime.Caption);
  ExamSeq := StrToIntDef(pnSeq.Caption,0);

  //sOpId:= DM.GetOpId(gdMaster.Cells[C_User, gdMaster.Row]);

  if Length(sOpId) <> 6 then begin
      ShowMessage('?????? USER ?????? ??????!');
      exit;
  end;

  sLoc := gdMaster.Cells[C_Locat, gdMaster.Row];
  sFlag:= gdMaster.Cells[C_Flag, gdMaster.Row];

end;

procedure TF_Find.gdMasterDblClick(Sender: TObject);
begin
  if gdMaster.Row > 0 then begin
      pnData.Visible:= True;
      edSendSpcid.SetFocus;
  end;
end;

procedure TF_Find.btnClosePanelClick(Sender: TObject);
begin
  pnData.Visible:= False;
end;

procedure TF_Find.edFlagChange(Sender: TObject);
begin
  pnOrdCode.Caption:= TPCode.GetPanelCode(edLocaion.Text, edFlag.Text);
end;

procedure TF_Find.edSendSpcidKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      SelectNext(Sender as TWinControl, True, True);
  end;
end;

procedure TF_Find.edSendPatIdKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      SelectNext(Sender as TWinControl, True, True);
  end;
end;

procedure TF_Find.edLocaionKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      SelectNext(Sender as TWinControl, True, True);
  end;
end;

procedure TF_Find.edFlagKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      SelectNext(Sender as TWinControl, True, True);
  end;
end;

procedure TF_Find.edUIdKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      SelectNext(Sender as TWinControl, True, True);
  end;
end;

procedure TF_Find.edSendPatIdKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_UP then
    SelectNext(Sender As TWinControl, False, True)
  else
    if Key = VK_DOWN then
      SelectNext(Sender As TWinControl, True, True);
end;

procedure TF_Find.edLocaionKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_UP then
    SelectNext(Sender As TWinControl, False, True)
  else
    if Key = VK_DOWN then
      SelectNext(Sender As TWinControl, True, True);
end;

procedure TF_Find.edFlagKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_UP then
    SelectNext(Sender As TWinControl, False, True)
  else
    if Key = VK_DOWN then
      SelectNext(Sender As TWinControl, True, True);
end;

procedure TF_Find.edUIdKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_UP then
    SelectNext(Sender As TWinControl, False, True)
  else
    if Key = VK_DOWN then
      SelectNext(Sender As TWinControl, True, True);
end;

procedure TF_Find.edSendSpcidKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_DOWN then
      SelectNext(Sender As TWinControl, True, True);
end;

procedure TF_Find.pnDataMouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
  ReleaseCapture;
  SendMessage(pnData.Handle, WM_SYSCOMMAND, 61458, 0);
end;

procedure TF_Find.btnSend_GridClick(Sender: TObject);
var
  ExamDate:string;
  ExamSeq:integer;
  i,iCount,iIndex:integer;
  bCheck:boolean;
  sBarCode, sOpId, sLoc, sFlag:string;
  nCount:integer;
  QCYN,POCTYN:boolean;
  sPatNo:string;
  TMaster: TIfMaster;
begin
  //TMaster:= TIfMaster.Create;

{
  iCount:=0;
  for i:=1 to gdMaster.RowCount -1 do begin
      bCheck:= False;
      gdMaster.GetCheckBoxState(C_Check, i, bCheck);
      if bCheck = True then begin
          iCount:= 1;
          Break;
      end;
  end;

  nCount:=0;

  if iCount > 0 then begin
      if MessageDlg('???? ?????? ?????? ????????????? ',
         mtWarning, mbOKCancel,1 ) = mrOk then
      begin
          for i:= gdMaster.RowCount-1 downto 1 do begin
              bCheck:= False;
              gdMaster.GetCheckBoxState(C_Check, i, bCheck);
              if bCheck = True then begin
                  ExamDate:= GetGridDate(Trim(gdMaster.Cells[C_Date, i]));
                  ExamSeq := StrToIntDef(Trim(gdMaster.Cells[C_Seq, i]),0);
                  sBarCode:= Trim(gdMaster.Cells[C_Spcid, i]);
                  sOpId   := Trim(gdMaster.Cells[C_User, i]);
                  sLoc    := Trim(gdMaster.Cells[C_Locat, i]);
                  sFlag   := Trim(gdMaster.Cells[C_Flag, i]);

                  if gdMaster.Cells[C_IDStat, i] = 'Q' then
                      QCYN:= True
                  else
                      QCYN:= False;

                  //?????? ??????????.
                  TGlobal.IDState:= 'B';

                  if DM.ReSend(ExamDate, sBarCode, sOpId, sLoc, sFlag, ExamSeq, QCYN) then begin
                      gdMaster.Cells[C_State,i]:= 'C';

                      if TGlobal.PatNo <> '' then begin
                          DM.SavePatId(TGlobal.PatNo, ExamDate, ExamSeq);
                          gdMaster.Cells[C_PatID, i]:= TGlobal.PatNo;
                      end;
                      inc(nCount);
                  end;
              end;
          end;

          ShowMessage(IntToStr(nCount)+'?? ????!');
      end;
  end
  else begin
      ShowMessage('?????? ?????? ???? ?????? ??????!');
  end;
  }

end;

procedure TF_Find.SetDefaultPanelColor(clDef:TColor);
begin
  pnCheck.Color    := clDef;
  pnOrdCreate.Color:= clDef;
  pnExcept.Color   := clDef;
  pnUpload.Color   := clDef;
end;

procedure TF_Find.btnCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TF_Find.gdMasterGridHint(Sender: TObject; ARow, ACol: Integer;
  var hintstr: String);
var
  sState:string;
begin
  if ARow > 0 then begin
      if ACol = C_State then begin
          sState:= Copy(gdMaster.Cells[ACol,ARow],1,1);
          if sState <> '' then begin
              Case sState[1] of
                  'U': hintstr:= '?????????? ????!';
                  'X': hintstr:= '???????? ????!';
                  'F': hintstr:= '?????? ???? Flag';
                  'N': hintstr:= '?????? ????????!';
                  'C': hintstr:= '???????? ????!';
                  'P': hintstr:= '?????? ????????!';
              end;
          end;
      end;
  end;

end;

procedure TF_Find.edLotNoKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_UP then
    SelectNext(Sender As TWinControl, False, True)
  else
    if Key = VK_DOWN then
      SelectNext(Sender As TWinControl, True, True);
end;

procedure TF_Find.edLotNoKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
      Key:= #0;
      SelectNext(Sender as TWinControl, True, True);
  end;
end;

end.





