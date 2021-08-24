unit U_Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Buttons, Menus, OleCtrls, ComCtrls, ExtCtrls,
  U_IfClass, StdCtrls, Grids, BaseGrid, AdvGrid, CPort,DateUtils,math;

type
  TF_Main = class(TForm)
    pnButton: TPanel;
    Panel2: TPanel;
    StatusBar1: TStatusBar;
    MainMenu1: TMainMenu;
    btnView: TSpeedButton;
    btnClose: TSpeedButton;
    gdIf: TAdvStringGrid;
    M1: TMenuItem;
    N2: TMenuItem;
    N1: TMenuItem;
    pnLog: TPanel;
    mmTemp: TMemo;
    mmLog: TMemo;
    pnInfo: TPanel;
    N3: TMenuItem;
    Debug1: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    btnRcv: TButton;
    N6: TMenuItem;
    N7: TMenuItem;
    btnUpload: TSpeedButton;
    dtpF: TDateTimePicker;
    Panel1: TPanel;
    PopupMenu1: TPopupMenu;
    N8: TMenuItem;
    dtpT: TDateTimePicker;
    Panel3: TPanel;
    cmbxOp: TComboBox;
    ComPort1: TComPort;
    Memo1: TMemo;
    btnClear: TSpeedButton;
    N9: TMenuItem;
    mnAuto: TMenuItem;
    N10: TMenuItem;
    N11: TMenuItem;
    DB1: TMenuItem;
    N21: TMenuItem;
    btnWork: TSpeedButton;
    procedure FormShow(Sender: TObject);
    procedure gdIfGetAlignment(Sender: TObject; ARow, ACol: Integer;
      var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure btnCloseClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure N2Click(Sender: TObject);
    procedure ComPort1RxChar(Sender: TObject; Count: Integer);
    procedure N5Click(Sender: TObject);
    procedure Debug1Click(Sender: TObject);
    procedure btnRcvClick(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure gdIfCanEditCell(Sender: TObject; ARow, ACol: Integer;
      var CanEdit: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure N7Click(Sender: TObject);
    procedure btnUploadClick(Sender: TObject);
    procedure btnViewClick(Sender: TObject);
    procedure btnClearClick(Sender: TObject);
    procedure mnAutoClick(Sender: TObject);
    procedure N10Click(Sender: TObject);
    procedure N11Click(Sender: TObject);
    procedure gdIfGetCellColor(Sender: TObject; ARow, ACol: Integer;
      AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure DB1Click(Sender: TObject);
    procedure gdIfClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure N21Click(Sender: TObject);
    procedure btnWorkClick(Sender: TObject);
  private
    { Private declarations }
    function PortOpen: boolean;

    procedure doChorus_Trio(aRcv: array of byte; nCnt:integer);
    procedure SendChorus_ACK;
    function ByteArrayToStr(aRcv: array of Byte):string;
    procedure InitGlobal;

    procedure Chorus_Trio_Rcv;
    procedure Chorus_Trio_ResProcess(sData:string);
    procedure SaveChorus_QC(TMaster:TIfMaster);
    function GetChorusRemark(IFCD, RES, FLG:string):string;
    function GetChorusResult(IFCD, RES, FLG:string):string;
    function GetChorusFlag(RES, FLG:string):string;

    procedure DoAlegria(const cData:string);
    procedure Alegria_Rcv;
    function GetAlegriaResText(IFCD, RES, FLG:string):string;

    procedure DoTest1(sData: string);
    procedure Test1_Process(sData:string);
    procedure Test1_OrderProcess(Barcode:string);
    procedure Test1_ResultProcess(sData:string);
    function Test1CheckSum(sData:string):string;
    procedure Test1_SendNoOrder(Barcode:string);

    procedure doEPOC_BT(const cData:string);
    procedure EPOC_BT_Rcv;

    procedure douTas(const cData:string);
    procedure uTas_Rcv;
    procedure uTas_OrderProcess(TMaster:TIfMaster);
    function Make_uTas_OrdStr(TMaster:TIfMaster):string;

    procedure doNsPlus(const cData:string);
    procedure NsPlus_Rcv;
    procedure NsPlus_Result_Process(sData:string);
    procedure NsPlus_Order_Process(sData:string);
    function Make_NsPlus_OrdStr(TMaster:TIfMaster):string;
    procedure NsPlus_Order_Send;

    procedure doG7(const cData:string);
    procedure G7_Rcv;

    procedure Rcv_Log(sData:string);
    procedure AddDebugLine(cStr:string);
    procedure SendStr(const cStr:string);
    procedure SetCommData(const cStr: string);
    procedure ClearCommData;
    procedure SendResultGet(sFrom, sTo, sBarCode:string);

    procedure DispOneResult(TMaster:TIfMaster; ARow:integer=0);
    function GetGridIndex(TGrid:TAdvStringGrid; ExamDate, ExamSeq:string; Add:boolean=true):integer;
    function AddOrdIndex:integer;
    function GetSampleDateTime(cData:string):TDateTime;
    procedure DispOnePat_Idx(TMaster:TIfMaster; GrdIdx:integer=0);
    procedure DispOnePat_Idx_Ord(TMaster:TIfMaster; GrdIdx:integer=0);
    procedure DispOnePat(TMaster:TIfMaster);
    procedure DispOnePat_ALL(TMaster:TIfMaster; GrdIdx:integer);
    procedure DispOnePat_ORD(TMaster:TIfMaster; GrdIdx:integer=0);
    procedure DispOneState(ExamDate, ExamSeq, ErrMsg, UpState:string);
    procedure DispOneState_IDX(GrdIdx:integer; UpState:string);

    function GetBCDIndex(TMaster:TIfMaster):integer;
    function GetAddIndex(TGrid:TAdvStringGrid):integer;

    function UploadOneExam(TObj: TObject; GrdIdx:integer; UploadHosp:boolean=True):boolean;
    procedure UploadProcess(TObj:TObject; GrdIdx:integer; IsAll:boolean=True);
    procedure UploadProcess_One(TObj:TObject; GrdIdx:integer);

    procedure MyFormCaptionChange(Check: boolean);
  public
    { Public declarations }
  end;

var
  F_Main: TF_Main;
  gACK:integer;
  vcTxData:string;
  slSender:TStringList;

  slRcv:TStringList;
  gTxbuffer: array [0..99] of Byte;    //전역Buffer
  gInputCount: Integer;                //총 받은갯수
  gNeedCount: Integer;      //장비에서 보내겠다고 알려준 데이터 갯수
  gCMDCode: Byte;           //장비 Comment Code

implementation
uses GlobalVar, U_CodeInfo, SetDataBase, U_DM, StringLib, U_CodeM,
  U_Server, U_Work;

const
  C_CKB = 0;
  C_DTM = C_CKB+1;
  C_BCD = C_DTM+1;
  C_PID = C_BCD+1;
  C_PNM = C_PID+1;
  C_POS = C_PNM+1;
  C_STA = C_POS+1;
  C_EDT = C_STA+1;
  C_SEQ = C_EDT+1;

{$R *.dfm}


{ TForm1 }

procedure TF_Main.FormShow(Sender: TObject);
var
  i,K:integer;
  iNm:string;
begin
  MyFormCaptionChange(mnAuto.Checked);

  iNM:= UpperCase(TGlobal.FIName);
  K:= C_SEQ;

  for i:=0 to TCode.TAbbr.Count -1 do begin
      INC(K);
      gdIf.InsertCols(K,1);
      if iNM = 'GEMINI' then
          gdIf.ColWidths[K]:= 120
      else
      if iNM = 'EPOC' then
          gdIf.ColWidths[K]:= 45
      else
          gdIf.ColWidths[K]:= 60;
      gdIf.Cells[K,0]:= TCode.TAbbr.Strings[i];
  end;

  gdIf.HideColumn(C_POS);

  gdIf.HideColumns(C_EDT, C_SEQ);

  Self.Top:= TGlobal.MainTop;
  Self.Left:= TGlobal.MainLeft;

  ComPort1.LoadSettings(stIniFile, TGlobal.AppPath+IniFileName);
  if not PortOpen then
      pnInfo.Color:= clRed;
end;

procedure TF_Main.DispOneState(ExamDate, ExamSeq, ErrMsg, UpState:string);
var
  nRow:integer;
begin
  nRow:= GetGridIndex(gdIf, ExamDate, ExamSeq, False);

  if nRow <= 0 then exit;

  if UpState = 'Y' then
      gdIf.Cells[C_STA, nRow]:= '전송완료'
  else
  if UpState = 'X' then
      gdIf.Cells[C_STA, nRow]:= '결과이상'
  else
      gdIf.Cells[C_STA, nRow]:= '미전송';

  StatusBar1.Panels[0].Text:= ErrMsg;

end;

procedure TF_Main.DispOneState_IDX(GrdIdx:integer; UpState:string);
var
  nRow:integer;
begin
  if (GrdIdx >= gdIf.RowCount) or (grdIdx<1) then exit;

  nRow:= GrdIdx;

  if UpState = 'Y' then
      gdIf.Cells[C_STA, nRow]:= '전송완료'
  else
  if UpState = 'X' then
      gdIf.Cells[C_STA, nRow]:= '결과이상'
  else
      gdIf.Cells[C_STA, nRow]:= '미전송';

end;


procedure TF_Main.gdIfGetAlignment(Sender: TObject; ARow, ACol: Integer;
  var HAlign: TAlignment; var VAlign: TVAlignment);
begin
  VAlign:= vtaCenter;
  HAlign:= taCenter;
end;

procedure TF_Main.btnCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TF_Main.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  if MessageDlg('종료하시겠습니까? ', mtWarning, mbOKCancel, 0) <> mrOk then
      CanClose:= False
  else
      CanClose:= True;
end;

procedure TF_Main.DispOneResult(TMaster:TIfMaster; ARow:integer=0);
var
  nRow, nCol: integer;
begin
  if (ARow >= gdIf.RowCount) or (ARow<1) then exit;

  nRow:= ARow;
  nCol:= GetAbbrIndex(gdIf, TMaster.FAbbr);

  if (nRow = 0) or (nCol = 0) then exit;

  with gdIf do begin
      Cells[nCol, nRow]:= TMaster.FResult;
      if TMaster.FFlag <> 'N' then
          gdIf.FontColors[nCol, nRow]:= clRed;
      //DP체크..
  end;

end;

procedure TF_Main.N2Click(Sender: TObject);
begin
  if F_CodeM = nil then
      F_CodeM:= TF_CodeM.Create(Self);

  F_CodeM.Show;
end;

procedure TF_Main.ComPort1RxChar(Sender: TObject; Count: Integer);
var
  cStr:string;
  iNm:string;
  a:array[0..99] of byte;
begin
  //ComPort1.ReadStr(cStr, Count);
  //Memo1.Lines.Add(cStr);
  iNm:= UpperCase(TGlobal.FIName);

  if INm = 'CHORUSTRIO' then begin
      FillChar(a, SizeOf(a), 0);

      ComPort1.Read(a, count);
      doChorus_Trio(a, count);
  end
  else begin
      ComPort1.ReadStr(cStr, Count);
      if INm = 'TEST1' then
          DoTest1(cStr)
      else
      if INm = 'ALEGRIA' then
          DoAlegria(cStr)
      else
      if INm = 'EPOC' then
          doEPOC_BT(cStr)
      else
      if INm = 'G7' then
          doG7(cStr)
      else
      if iNm = 'UTAS' then
          doUTAS(cStr);
  end;

end;

procedure TF_Main.Rcv_Log(sData: string);
var
  F:TextFile;
begin
  if Not FileExists(DataFileName) then
      exit;

  try
      AssignFile(F,DataFileName);
      Append(F);
      Writeln(F, FormatDateTime('yyyy-mm-dd hh:nn:ss', now) + #13#10 + sData);
  finally
      CloseFile(F);
  end;
end;

procedure TF_Main.AddDebugLine(cStr: string);
begin
  if pnLog.Visible = True then
      mmLog.Lines.Add(cStr);
end;

procedure TF_Main.SendStr(const cStr: string);
begin
  if cStr = ENQ then
      mmLog.Lines.Add('HOST ->[ENQ]')
  else
  if cStr = ACK then
      mmLog.Lines.Add('HOST ->[ACK]')
  else
  if cStr = EOT then
      mmLog.Lines.Add('HOST ->[EOT]')
  else
      mmLog.Lines.Add('Host ->'+cStr);

  //delay(100);
  sleep(200);
  ComPort1.WriteStr(cStr);

end;

procedure TF_Main.SetCommData(const cStr: string);
begin
    if TGlobal.FIName = 'EPOC' then
      mmTemp.Lines.Text:= cStr
  else
      mmTemp.Text:= cStr;

  TGlobal.DataLog:= mmTemp.Text;
end;

procedure TF_Main.ClearCommData;
begin
  mmTemp.Text:= '';
end;

procedure TF_Main.DispOnePat(TMaster: TIfMaster);
var
  nRow:integer;
begin

  nRow:= GetGridIndex(gdIf, TMaster.FExamDate, TMaster.FExamSeq);

  gdIf.AddCheckBox(0, nRow, False, False);
  gdIf.Cells[C_DTM, nRow] := TMaster.FRcvTime;
  gdIf.Cells[C_SEQ, nRow] := TMaster.FExamSeq;
  gdIf.Cells[C_EDT, nRow] := TMaster.FExamDate;
  gdIf.Cells[C_BCD, nRow] := TMaster.FBarCode;
  gdIf.Cells[C_PID, nRow] := TMaster.FPId;
  gdIf.Cells[C_PNM, nRow] := TMaster.FPNm;

  if TMaster.FRack <> '' then
      gdIf.Cells[C_POS, nRow] := TMaster.FRack +'-'+ TMaster.FPos
  else
      gdIf.Cells[C_POS, nRow] := TMaster.FPos;

  //gdIf.Cells[C_ABR, nRow] := TMaster.FAbbr;
  //gdIf.Cells[C_RES, nRow] := TMaster.FResult;

  if TMaster.FUpState = 'U' then
      gdIf.Cells[C_STA, nRow] := '전송완료'
  else
  if TMaster.FUpState = 'Y' then
      gdIf.Cells[C_STA, nRow] := '전송완료'
  else
  if TMaster.FUpState = 'X' then
      gdIf.Cells[C_STA, nRow] := '결과이상'
  else begin
      if TMaster.FOrdState = 'N' then
          gdIf.Cells[C_STA, nRow] := '오더없음';

  end;

  gdIf.AutoNumberCol(0);

end;

function TF_Main.GetGridIndex(TGrid: TAdvStringGrid; ExamDate,
  ExamSeq: string; Add: boolean): integer;
var
  i:integer;
  cDateTime:string;
  cSeq:string;
begin
  Result:= 0;
  if (TGrid.RowCount = 2) and (TGrid.Cells[C_DTM,1]='') then begin
      Result:=1;  exit;
  end;

  for i:=TGrid.RowCount -1 downto 1 do begin
      if (TGrid.Cells[C_EDT, i] = ExamDate) and (TGrid.Cells[C_SEQ, i] = ExamSeq) then
      begin
          Result:= i;
          exit;
      end;
  end;

  if Add = True then begin
      TGrid.AddRow;
      Result:= TGrid.RowCount -1;
  end;
end;

procedure TF_Main.SendResultGet(sFrom, sTo, sBarCode: string);
var
  sHeader,sQuery,sEnd:string;
  dStrTm:TDateTime;
  HCS, MCS, ECS:string;
begin

  slSender.Clear;

  sHeader:= '1H|\^&|||1234567890|||||||P||'+FormatDateTime('yyyymmddhhmmdd', now)+CR+ETX;
  HCS:=ASTMCheckSum(sHeader);
  slSender.Add( STX+sHeader+HCS+CR+LF );

  //if sBarCode = '' then
  //    sQuery:= '2Q|1|99999||||'+sFrom+'|'+sTo+'|||S||F'+CR+ETX
  //else
  sQuery:= '2Q|1|'+sBarCode+'||||'+sFrom+'|'+sTo+'|||S||F'+CR+ETX;

  MCS:= ASTMCheckSum(sQuery);
  slSender.Add( STX+sQuery+MCS+CR+LF );

  sEnd:='3L|1|N|'+CR+ETX;
  ECS:= ASTMCheckSum(sEnd);
  slSender.Add( STX+sEnd+ECS+CR+LF );

  gACK:=0;

  SendStr(ENQ);

end;

function TF_Main.PortOpen: boolean;
begin
  Result:= False;
  try
      ComPort1.Open;
  except
      on e:exception do
      begin
          Application.MessageBox(PChar(e.Message), 'Port Error!');
          exit;
      end;
  end;

  Result:= True;
end;

procedure TF_Main.N5Click(Sender: TObject);
begin
  Close;
end;

procedure TF_Main.Debug1Click(Sender: TObject);
begin
  pnLog.Visible:= debug1.Checked;
  btnRcv.Visible:= Debug1.Checked;
end;

procedure TF_Main.btnRcvClick(Sender: TObject);
var
  iNm:string;
begin
  iNm:= UpperCase(TGlobal.FIName);
  if iNm = 'CHORUSTRIO' then begin
      slRcv.Text:= mmTemp.Text;
      Chorus_Trio_Rcv;
  end
  else
  if iNm = 'G7' then
      G7_Rcv
  else
  if INm = 'ALEGRIA' then
      Alegria_Rcv
   else
  if INm = 'EPOC' then
      EPOC_BT_Rcv
  else
  if iNm = 'TEST1' then
      Test1_Process(mmTemp.Text)
  else
  if iNm = 'UTAS' then
      uTas_Rcv;
end;

procedure TF_Main.N6Click(Sender: TObject);
begin
  ComPort1.ShowSetupDialog;
  ComPort1.StoreSettings(stIniFile, TGlobal.AppPath + IniFileName);
end;

procedure TF_Main.gdIfCanEditCell(Sender: TObject; ARow, ACol: Integer;
  var CanEdit: Boolean);
begin
  if ARow > 0 then begin
      if ACol in [C_CKB, C_BCD] then
          CanEdit:= True
      else
          CanEdit:= False;
  end;
end;

procedure TF_Main.FormCreate(Sender: TObject);
begin
  dtpF.Date:= now;
  dtpT.Date:= now;

  pnInfo.Caption:= TGlobal.FTitle;
  mnAuto.Checked:= TGlobal.FAutoSend;

  Application.Title:= TGlobal.FTitle;
  InitGlobal;

  Self.Top:= TGlobal.MainTop;
  Self.Left:= TGlobal.MainLeft;
  Self.Height:= TGlobal.MainHeight;
  Self.Width:= TGlobal.MainWidth;
end;

procedure TF_Main.N7Click(Sender: TObject);
begin
  if MessageDlg('화면상의 결과 데이터를 삭제하시겠습니까?', mtConfirmation, mbOKCancel, 1) = mrOk then begin
      gdIf.ClearNormalCells;
      gdIf.RowCount:=2;
  end;
end;

procedure TF_Main.btnUploadClick(Sender: TObject);
var
  TMaster:TIfMaster;
  i, iCount, UpCount:integer;
  bCheck:boolean;
begin

  iCount:=0;  UpCount:= 0;
  for i:=1 to gdIf.RowCount -1 do begin
      bCheck:= False;
      gdIf.GetCheckBoxState(0, i, bCheck);
      if bCheck = True then begin
          iCount:= 1;
          Break;
      end;
  end;

  if iCount > 0 then begin
      if MessageDlg('선택 결과를 재전송 하시겠습니까? ',
         mtWarning, mbOKCancel,1 ) = mrOk then
      begin
          for i:= 1 to gdIf.RowCount -1 do begin
              bCheck:= False;
              gdIf.GetCheckBoxState(0, i, bCheck);
              if bCheck = True then begin
                  TMaster:= TIfMaster.Create;
                  try
                      TMaster.FOrdState:= 'N';
                      TMaster.FExamDate:= gdIf.Cells[C_EDT, i];
                      TMaster.FExamSeq := gdIf.Cells[C_SEQ, i];
                      TMaster.FBarCode := gdIf.Cells[C_BCD, i];

                      DM.DownLoadOrder_DJI(TMaster);

                      if Debug1.Checked then
                          ShowMessage(TMaster.FExamDate+','+TMaster.FExamSeq+','+TMaster.FBarCode+','+TMaster.FOrdState);

                      if TMaster.FOrdState = 'Y' then begin
                          DM.SaveMaster(TMaster);
                          gdIf.Cells[C_PID, i]:= TMaster.FPID;
                          gdIf.Cells[C_PNM, i]:= TMaster.FPNM;

                          if DM.UploadResult(TMaster) then
                          begin
                              Inc(UpCount);
                              gdIf.SetCheckBoxState(0, i, False);
                              DM.SaveState(TMaster.FExamDate, TMaster.FExamSeq, TGlobal.ErrMsg, 'Y');
                              DispOneState_IDX(i, 'Y');
                          end
                          else begin
                              DispOneState_IDX(i, 'N');
                          end;
                      end
                      else begin
                          gdIf.Cells[C_STA, i]:= '오더없음';
                      end;
                  finally
                      TMaster.Free;
                  end;
              end;
          end;

          ShowMessage(IntToStr(UpCount)+'건 전송!');
      end;
  end
  else begin
      ShowMessage('전송할 검사를 선택 하셔야 합니다!');
  end;

  gdIf.Refresh;

end;

procedure TF_Main.btnViewClick(Sender: TObject);
var
  TMaster:TIfMaster;
  GrdIdx:integer;
  OldSeq:string; NewSeQ:string;
begin
  gdIf.Row:=1;
  gdIf.ClearNormalCells;
  gdIf.RowCount:=2;

  GrdIdx:=0;

  with DM.qryV do begin
      Close;
      SQL.Text:= ' Select M.ExamDate, M.BarCode, M.ExamTime, M.ExamSeq, M.PID, M.RPOS, M.PNM, M.ANO, '+
                 '        M.Age, M.Sex, M.OrdCode, R.PName, R.examcode, M.UpState, R.UpCode, R.Result, R.Flag   '+
                 ' From TB_Master M Inner Join TB_Result R '+
                 ' on (M.ExamDate = R.ExamDate  And        '+
                 '     M.ExamSeq  = R.ExamSeq )            '+
                 ' Where Mid(M.ExamTime,1,10) Between '''+FormatDateTime('yyyy-mm-dd', dtpF.Date)+''' '+
                 '                                and '''+FormatDateTime('yyyy-mm-dd', dtpT.Date)+''' ';
      if cmbxOp.ItemIndex > 0 then begin
          if cmbxOp.ItemIndex = 2 then
              SQL.Text:= SQL.Text + ' And M.UpState = ''Y'' '
          else
              SQL.Text:= SQL.Text + ' And M.UpState <> ''Y'' ';
      end;
      SQL.Text:= SQL.Text + '   Order By M.ExamDate, M.ExamSeq ';
      Open;

      if RecordCount = 0 then exit;

      OldSeq:='';
      NewSeq:='';
      while Not Eof do begin
          NewSeq:= FieldByName('ExamSeq').AsString;

          TMaster:= TIfMaster.Create;
          try
              TMaster.FExamDate:= FieldByName('ExamDate').AsString;
              TMaster.FExamSeq := FieldByName('ExamSeq').AsString;
              TMaster.FRcvTime := FieldByName('ExamTime').AsString;
              TMaster.FOrdCode  := FieldByName('OrdCode').AsString;
              TMaster.FPId := FieldByName('PID').AsString;
              TMaster.FPNm := FieldByName('PNM').AsString;
              TMaster.FAge   := FieldByName('Age').AsString;
              TMaster.FSex   := FieldByName('Sex').AsString;
              TMaster.FUpState:= FieldByName('UpState').AsString;
              TMaster.FExamPanel:= FieldByName('PName').AsString;
              TMaster.FBarCode:= FieldByName('BarCode').AsString;
              TMaster.FRack   := TokenStr(FieldByName('RPOS').AsString,'-',1);
              TMaster.FPos    := TokenStr(FieldByName('RPOS').AsString,'-',2);
              TMaster.FANO    := Copy(FieldByName('ANO').AsString,3,10);

               if OldSeq <> NewSeq then begin
                  GrdIdx:= GetAddIndex(gdIf);
                  DispOnePat_Idx(TMaster, GrdIdx);
                  OldSeq:= NewSeq;
              end;

              TMaster.FExamCode:= FieldByName('examcode').AsString;
              TMaster.FUpCode  := FieldByName('UpCode').AsString;
              TMaster.FAbbr  := TCode.GetAbbr_Up(TMaster.FUpCode);
              TMaster.FResult:= FieldByName('Result').AsString;
              TMaster.FFlag  := FieldByName('Flag').AsString;

              DispOneResult(TMaster, GrdIdx);

              //DispOnePat_ALL(TMaster, GrdIdx);
          finally
              TMaster.Free;
          end;

          Next;
      end;
  end;

  gdIf.Refresh;

end;

function TF_Main.GetBCDIndex(TMaster: TIfMaster): integer;
var
  bFind:boolean;
  i:integer;
begin
  Result:= 0;
  bFind:= False;

  if (gdIf.RowCount =2 ) and (gdIf.Cells[C_EDT,1]='') then
  begin
      Result:=1;
      exit;
  end;

  for i:=gdIf.RowCount -1 downto 1 do begin
      if TMaster.FBarCode = gdIf.Cells[C_BCD,i] then begin
          Result:= i;
          TMaster.FExamSeq:= gdIf.Cells[C_SEQ, i];
          bFind:= True;
          exit;
      end;
  end;

  if bFind = False then begin
      gdIf.AddRow;
      Result:= gdIf.RowCount -1;
  end;


end;

procedure TF_Main.DispOnePat_Idx(TMaster: TIfMaster; GrdIdx: integer);
var
  ExamTime:string;
  nRow:integer;
begin
  if (GrdIdx >= gdIf.RowCount) or (grdIdx<1) then exit;

  nRow:= GrdIdx;

  ExamTime:= FormatDateTime('yyyy-mm-dd hh:nn:ss', GetSampleDateTime(TMaster.FInstTime));

  gdIf.AddCheckBox(0, nRow, True, False);
  gdIf.Cells[C_DTM, nRow] := ExamTime;

  gdIf.Cells[C_EDT, nRow] := TMaster.FExamDate;
  gdIf.Cells[C_SEQ, nRow] := TMaster.FExamSeq;
  gdIf.Cells[C_BCD, nRow] := TMaster.FBarCode;
  //gdIf.Cells[C_ABR, nRow] := TMaster.FAbbr;
  gdIf.Cells[C_PID, nRow] := TMaster.FPID;
  gdIf.Cells[C_PNM, nRow] := TMaster.FPNM;
  if TMaster.FRack <> '' then
      gdIf.Cells[C_POS, nRow] := TMaster.FRack+'-'+ TMaster.FPos
  else
      gdIf.Cells[C_POS, nRow] := TMaster.FPos;

  if TMaster.FUpState = 'Y' then begin
      gdIf.Cells[C_STA, nRow] := '전송완료'
  end
  else begin
      if TMaster.FUpState = 'X' then
          gdIf.Cells[C_STA, nRow] := '결과이상'
      else
          gdIf.Cells[C_STA, nRow] := '미전송';
  end;

  gdIf.AutoNumberCol(0);

end;

function TF_Main.GetSampleDateTime(cData: string): TDateTime;
var
  yyyy,mm,dd,hh,nn,ss:string;
  Date,Time:string;
begin
  Result:= now;
  //20080326112638
  if Length(cData) <> 14 then exit;

  yyyy:= Copy(cData,1,4);
  mm  := Copy(cData,5,2);
  dd  := Copy(cData,7,2);
  hh  := Copy(cData,9,2);
  nn  := Copy(cData,11,2);
  ss  := Copy(cData,13,2);

  Date:= yyyy + '-' + mm + '-' + dd;
  Time:= hh + ':' + nn + ':' + ss;

  Result:= StrToDateTimeDef(Date+' '+Time, now);


end;

procedure TF_Main.doChorus_Trio(aRcv: array of byte; nCnt:integer);
var
  i: integer;
  Str: string;
begin
  try
  for i:=0 to nCnt -1 do begin
      Inc(gInputCount);                              //총 받은갯수
      gTxbuffer[gInputCount - 1]:= aRcv[i];

      Case gInputCount of
          2: gNeedCount:= gTxbuffer[gInputCount - 1]; //장비에서 보내겠다고 알려준 데이터 갯수
          3: gCMDCode:= gTxbuffer[gInputCount - 1];   //장비 Comment Code
          else begin
              if( gInputCount = gNeedCount + 3 ) then  begin //gNeedCount + STX + CS
                    Case gCMDCode of
                        5: SendChorus_ACK;                        //ENQ - 데이터 START
                        210: ;  //SendChorusOrder;              //D2 - 오더요청
                        211: SendChorus_ACK;                      //D3 - 데이터 END

                        215: begin                         //D7 - 결과데이터, 한건씩 처리하려면 여기서!!
                                 Str:= ByteArrayToStr(gTxBuffer);
                                 slRcv.Add(Str);
                                 //Chorus_Trio_Rcv(Str);
                                 SendChorus_ACK;
                             end;
                        216: begin
                                 //mmTemp.Lines.Add(Str);
                                 SendChorus_ACK;
                                 SetCommData(slRcv.Text);
                                 //mmTemp.Text:= slRcv.Text;
                                 Chorus_Trio_Rcv;       //D8 - 결과데이터 END, 결과를 모두 받은후에 한번에 처리하려면 여기서 해야함!!
                                 slRcv.Clear;
                             end;
                    end;

                    InitGlobal;
                                              //데이터 다 나와서 처리했으니 초기화
                End;
          end;
      end;
  end;
  except
      on e:exception do ShowMessage(e.Message);
  end;
end;

procedure TF_Main.btnClearClick(Sender: TObject);
begin
  gdIf.ClearNormalCells;
  gdIf.Row:=1;
  gdIf.RowCount:=2;
  gdIf.Refresh;
end;

procedure TF_Main.doNsPlus(const cData: string);
var
  i:integer;
  cStr,cBuffer:string;
  sTmp:string;
  iNm:string;
begin
  iNm:= TGlobal.FIName;

   // 전역변수로 보관한 자료를 읽음
   cStr:='';
   cBuffer:=cData;
   for i:= 1 to Length(cBuffer) do begin
     case cBuffer[i] of
     ACK:begin
             AddDebugLine(iNm+'->[ACK]');
             NsPlus_Order_Send;
             sTmp:='';
         end;
     ENQ:begin
             AddDebugLine(iNm+'->[ENQ]');
             SendStr(ACK); //ENQ
             vcTxData:='';
             sTmp:='';
         end;
     STX:begin
             //sTmp:= vcTxData+'[STX]';
             vcTxData:= STX;
         end;
     NAK:begin
             AddDebugLine(iNm+'->[NAK]');
         end;
     ETX:begin
             //sTmp:= vcTxData+'[ETX]';
             //sTmp:='';
             vcTxData:= vcTxData+ETX;
             SendStr(ACK);
         end;
     EOT:begin
             sTmp:= vcTxData+'[EOT]';
             AddDebugLine(iNm+'->'+sTmp);
             try
                 SetCommData(vcTxData);
                 NsPlus_Rcv;
             finally
                 sTmp:='';
                 vcTxData:='';
                 //ClearCommData;
             end;
         end;
         else begin
            vcTxData:=vcTxData+cBuffer[i];
         end;
     end;
   end;

end;

procedure TF_Main.NsPlus_Order_Process(sData: string);
var
  TMaster:TIfMaster;
  grdIdx:integer;
  sOrder, sOrd:string;
begin
  slSender.Clear;

  if Trim( sData ) = '' then exit;

  TMaster:= TIfMaster.Create;
  try
      TMaster.FBarCode := Trim(Copy(sData, 3, 14));
      TMaster.FRack    := Trim(Copy(sData, 17, 3));
      TMaster.FPos     := Trim(Copy(sData, 21, 2));

      //Exec_KSM_Reg(TMaster);

      //DownLoadOrder_KSM_Order(TMaster);

      grdIdx:= GetAddIndex(gdIf);

      sOrd:= Make_NsPlus_OrdStr(TMaster);

      DispOnePat_Idx(TMaster, GrdIdx);

      {
        Text section   1 	W 		   Type of text
        Rack ID		   6	001-05		 Rack ID position (to be filled from the left)
        ID 		       14 12345678	 Sample ID（to be filled from the left）
        Meas. No.	   1	2		       Number of measurements
        Item 		     2 	1		       Protocol No.（to be filled from the left）
                     3	1		       Dilution factor（to be filled from the right） }

      {sOrder:=  '2   14   1';
      TMaster.FOrdCnt:= 2;
      sOrder:= STX+ 'W      '+PadRightStr(TMaster.FBarCode,' ',14)+IntToStr(TMaster.FOrdCnt) + sOrder + ETX;
      slSender.Text:= sOrder + ASTMCheckSum(sOrder);}

      if TMaster.FOrdCnt = 0 then
          sOrder:= STX+'W      '+PadRightStr(TMaster.FBarCode,' ',14)+'0'+ETX
      else
          sOrder:= 'W      '+PadRightStr(TMaster.FBarCode,' ',14)+IntToStr(TMaster.FOrdCnt) + sOrd + ETX;

      slSender.Text:= sOrder + ASTMCheckSum(sOrder);

      gACK:=0;

      SendStr(ENQ);
      
      gdIf.Refresh;

  finally
      TMaster.Free;
  end;

end;

procedure TF_Main.NsPlus_Rcv;
var
  i:integer;
  sData,cFrame:string;
begin
  Rcv_Log(mmTemp.Text);

  for i:=0 to mmTemp.Lines.Count -1 do begin
      sData:= mmTemp.Lines.Strings[i];
      if Trim( sData ) = '' then continue;

      cFrame:=Copy(sData, 2, 1);

      if cFrame='D' then
          NsPlus_Result_Process(sData)
      else
      if cFrame='R' then
          NsPlus_Order_Process(sData)
  end;
end;

procedure TF_Main.NsPlus_Result_Process(sData: string);
var
  TMaster:TIfMaster;
  SampleType, SampleSeq:string;
  GrdIdx:integer;
begin

  TMaster:= TIfMaster.Create;

  try // ~Finally
      if Trim( sData ) = '' then exit;

      SampleType:= Copy(sData, 2, 1);  //U:Sample, C:Control, S:Calibrator
      SampleSeq := Trim(Copy(sData, 3, 4));  //Sample:1, STAT:8001, Control:9001

      TMaster.FRack:= Trim(Copy(sData, 7, 3));
      TMaster.FPos := Trim(Copy(sData, 10, 2));
      //Copy(12,1) Meas.method 1 1		1:1day, 2:2day, 3:3day
      TMaster.FBarCode:= Trim(Copy(sData, 14, 13));
      TMaster.FIfCode := Trim(Copy(sData, 68, 2));
      TMaster.FResult := Trim(Copy(sData, 35, 4));
      TMaster.FExamDate:= FormatDateTime('yyyymmdd',now);//Copy(sData,45,4)+Copy(sData,50,2)+Copy(sData,53,2);
      TMaster.FInstTime:= TMaster.FExamDate + PadLeftStr(Trim(Copy(sData,55,2)),'0',2)+
                                                              Copy(sData,58,2)+
                                                              Copy(sData,61,2);
      TMaster.FReMark  := Copy(sData,64,2);
      if TMaster.FReMark = '00' then
          TMaster.FFlag:= 'N'
      else
          TMaster.FFlag:= TMaster.FReMark;

      GrdIdx:= GetBCDIndex(TMaster);

      if TMaster.FExamSeq = '' then begin
          TMaster.FExamSeq:= DM.GetExamSeq(TMaster.FExamDate);
          //DownLoadOrder_KSM_Result(TMaster);
      end;

      DispOnePat_Idx(TMaster, GrdIdx);
      DM.SaveMaster(TMaster);

      UploadOneExam(TMaster, GrdIdx);
      UploadProcess(TMaster, GrdIdx, True);

  finally
      TMaster.Free;
  end;

end;

procedure TF_Main.NsPlus_Order_Send;
begin
  Inc(gACK);

  case gACK of
      1: SendStr(slSender.Strings[0]);
      else begin
          gACK:=0;
          SendStr(EOT);
      end;
  end;

end;

function TF_Main.GetAddIndex(TGrid: TAdvStringGrid): integer;
begin
  Result:= 0;

  if (TGrid.RowCount = 2) and (TGrid.Cells[C_EDT, 1]='') then
  begin
      Result:= 1;
      exit;
  end
  else begin
      TGrid.AddRow;
      Result:= TGrid.RowCount -1;
  end;
end;

function TF_Main.Make_NsPlus_OrdStr(TMaster: TIfMaster): string;
var
  VC, i, nCnt:integer;
  sO, sOrder:string;
begin
  Result:= '';
  sOrder:= '';
  nCnt:=0;

  {
  Hemoglobin (A) 	1
	Hemoglobin (N)	2
	Hemoglobin (GC)	3
	Transferin (A)	4
	Transferin (N)	5
	Transferin (GC)	6

  Text section   1 	W 		   Type of text
  Rack ID		   6	001-05		 Rack ID position (to be filled from the left)
  ID 		       14 12345678	 Sample ID（to be filled from the left）
  Meas. No.	   1	2		       Number of measurements
  Item 		     2 	1		       Protocol No.（to be filled from the left）
               3	1		       Dilution factor（to be filled from the right）

  'W      '+TMaster.FBarCode+'2'+'1 '+'1  '
                                +'4 '+'1  ';   }

  VC:= VarArrayDimCount(TMaster.vOrder);
  if VC > 0 then begin
      for i:=0 to VarArrayHighBound(TMaster.vOrder,1) do begin
         sO:= Trim(TMaster.vOrder[i]);
         if sO <> '' then begin
             sOrder:= sOrder + PadRightStr(sO,' ',2) + '  1';
             Inc(nCnt);
         end;
      end;
  end;

  if sOrder <> '' then
      Result:= sOrder;

  TMaster.FOrdCnt:= nCnt;
end;

function TF_Main.UploadOneExam(TObj: TObject; GrdIdx: integer;
  UploadHosp: boolean): boolean;
var
  TMaster: TIfMaster absolute TObj;
  VCode:Variant;
  i, CodeCnt:integer;
begin
  Result:= False;

  //검사코드 리스트를 가져와서 처리한다.(여러검사코드 있는경우 vCode루프..)
  VCode:= TCode.GetExamCode_Var(TMaster.FIfCode);

  CodeCnt:= VarArrayHighBound(VCode, 1);
  for i:=0 to CodeCnt do begin
      TMaster.FExamCode:= VCode[i];
      if TMaster.FExamCode <> '' then begin

          //등록코드라면..
          if TCode.SetCode_IfCode(TMaster.FIfCode) then begin

              //그리드에 검사 세팅되있는것 처리..
              TMaster.FAbbr:= TCode.GetAbbr(TMaster.FExamCode);
              if TMaster.FAbbr <> '' then begin

                  //화면에표시, 결과저장
                  DispOneResult(TMaster, GrdIdx);

                  DM.SaveResult(TMaster);
              end;
          end;
      end;
  end;      //For iLoof

end;

procedure TF_Main.UploadProcess(TObj: TObject; GrdIdx:integer; IsAll: boolean);
var
  TMaster:TIfMaster absolute TObj;
begin
  if mnAuto.Checked = False then
      exit;

  if DM.UploadResult(TMaster) = True then begin
      TMaster.FUpState:= 'Y';
      DM.SaveState( TMaster.FExamDate, TMaster.FExamSeq, TGlobal.ErrMsg, TMaster.FUpState);
  end;

  DispOneState_IDX( GrdIdx, TMaster.FUpState);

end;

function TF_Main.GetChorusFlag(RES, FLG: string): string;
var
  dVal:double;
begin
  //if FLG = 'N' then
  //    Result:= 'N'
  //else begin
      dVal:= StrToFloatDef(RES, -1);
      if dVal > 0 then begin
          if dVal > 1.1 then
              Result:= 'P'
          else
          if (dVal <= 0.9) and (dVal >= 1.1) then
              Result:= 'B'
          else
          if (dVal < 0.9) then
              Result:= 'N';
      end
      else begin
          Result:= FLG;
      end;
  //end;

end;

procedure TF_Main.mnAutoClick(Sender: TObject);
begin
  TGlobal.FAutoSend:= mnAuto.Checked;
  MyFormCaptionChange(mnAuto.Checked);
end;

procedure TF_Main.MyFormCaptionChange(Check: boolean);
begin
  if mnAuto.Checked = true then
      Self.Caption:= ' SANSOFT '+TGlobal.FTitle+' Interface Program' + '    ◈ 결과 자동 전송 ◈'
  else
      Self.Caption:= ' SANSOFT '+TGlobal.FTitle+' Interface Program';

  if TGlobal.FUserID <> '' then
      Self.Caption:= Self.Caption + '  검사자: '+TGlobal.FUserID;
end;

procedure TF_Main.DispOnePat_Idx_Ord(TMaster: TIfMaster; GrdIdx: integer);
var
  nRow, i, j:integer;
  ExamTime, Abr:string;
begin
  if GrdIdx = 0 then begin
      nRow:= GetAddIndex(gdIf);
  end
  else begin
      nRow:= GrdIdx;
  end;

  ExamTime:= FormatDateTime('yyyy-mm-dd hh:nn:ss', GetSampleDateTime(TMaster.FInstTime));

  gdIf.AddCheckBox(0, nRow, False, False);
  gdIf.Cells[C_DTM, nRow] := ExamTime;
  gdIf.Cells[C_SEQ, nRow] := '';
  gdIf.Cells[C_EDT, nRow] := '';
  gdIf.Cells[C_BCD, nRow] := TMaster.FBarCode;
  gdIf.Cells[C_PID, nRow] := TMaster.FPID;
  gdIf.Cells[C_PNM, nRow] := TMaster.FPNM;
  gdIf.Cells[C_POS, nRow] := TMaster.FRack+TMaster.FPos;

  if TMaster.FOrdState = 'N' then
      gdIf.Cells[C_STA, nRow] := '오더없음'
  else
      gdIf.Cells[C_STA, nRow] := '오더전송';

  for i:= C_SEQ+1 to gdIf.AllColCount -1 do begin
      for j:=0 to TMaster.FOrdCnt -1 do begin
          if TMaster.vAbbr[j] = gdIf.Cells[i, 0] then
          begin
              gdIf.Cells[i, nRow]:= 'Y';
              Break;
          end;
      end;
  end;

  gdIf.AutoNumberCol(0);

end;

procedure TF_Main.douTas(const cData: string);
var
  i:integer;
  cStr,cBuffer:string;
  sTmp:string;
begin
   // 전역변수로 보관한 자료를 읽음
   cStr:='';
   cBuffer:=cData;
   for i:= 1 to Length(cBuffer) do begin
     case cBuffer[i] of
     ACK:begin
             AddDebugLine('uTAS->[ACK]');
             //SendMyData;
             sTmp:='';
         end;
     NAK:begin
             AddDebugLine('uTAS->[NAK]');
         end;
     ENQ:begin
             AddDebugLine('uTAS->[ENQ]');
             SendStr(ACK); //ENQ
             vcTxData:='';
             sTmp:='';
         end;
     STX:begin
             sTmp:= vcTxData+'[STX]';
             vcTxData:= STX;
         end;
     ETB:begin
             sTmp:= vcTxData+'[ETB]';
             vcTxData:= vcTxData + ETB;
             sTmp:='';
             //SendStr(ACK);
         end;
     ETX:begin
             sTmp:= vcTxData+'[ETX]';
             sTmp:='';
             SendStr(ACK);

             vcTxData:= vcTxData + ETX;
             SetCommData(vcTxData);
             uTas_Rcv;

             sTmp:='';
             vcTxData:='';
         end;
     LF :begin
             vcTxData:= vcTxData + LF;
         end;
     EOT:begin
         end;
         else begin
            vcTxData:=vcTxData+cBuffer[i];
         end;
     end;
   end;

end;

function TF_Main.Make_uTas_OrdStr(TMaster: TIfMaster): string;
var
  VC, i, nCnt:integer;
  sO, sOrder:string;
begin
  Result:= '00,';
  sOrder:= '';
  nCnt:=0;

  VC:= VarArrayDimCount(TMaster.vOrder);
  if VC > 0 then begin
      for i:=0 to VarArrayHighBound(TMaster.vOrder,1) do begin
         sO:= Trim(TMaster.vOrder[i]);
         if sO <> '' then begin
             if sO = '99' then begin
                 if Pos('01',sOrder) < 1 then begin
                     sOrder:= sOrder + '01' + ',';
                     Inc(nCnt);
                 end;
                 if Pos('02',sOrder) < 1 then begin
                     sOrder:= sOrder + '02' + ',';
                     Inc(nCnt);
                 end;
             end
             else begin
                 if Pos(sO,sOrder) < 1 then begin
                     sOrder:= sOrder + Trim(TMaster.vOrder[i]) + ',';
                     Inc(nCnt);
                 end;
             end;
         end;
      end;
  end;

  if sOrder <> '' then
      Result:= sOrder;

  TMaster.OrdCnt:= nCnt;

end;

procedure TF_Main.uTas_OrderProcess(TMaster: TIfMaster);
var
  sOrder, sSender, Bcd:string;
begin
  DM.DownLoadOrder_SCHUH(TMaster);

  sOrder:= Make_uTas_OrdStr(TMaster);
  //'K,   1210181412,     ,00001,1,03,99,02,01,                i'
  sSender:= 'K,' +
            TMaster.FBarCode+','+
            '     ,'+                                             //SampleNo
            '00001,'+                                             //Dilution
            '1,'    +                                             //SampleType = serum fix
            PadLeftStr(IntToStr(TMaster.OrdCnt),'0',2)+','+       //ItemCount
            sOrder+                                               //Item(,포함)
            '                '+ ETX;                              //Comment(16byte)

  sSender:= STX + sSender + BCCCheckSum_Char(sSender);

  TGlobal.DataLog:= 'IF_PC-> '+sSender;

  DispOnePat_ORD(TMaster);

  SendStr(sSender);

  gdIf.Refresh;

end;

procedure TF_Main.uTas_Rcv;
var
  TMaster:TIfMaster;
  i, j, GrdIdx, ResCnt:integer;
  bUpSucc:boolean;
  sData,sFrame,ResFlag:string;
  dAFP, dAfpL3, dRES:double;
  sAFP, sAfpL3:string;
begin

  GrdIdx:=0;
  TMaster:= TIfMaster.Create;
  dAfP:=0;  dAfpL3:=0;
  sAFP:=''; sAfpL3:= '';

  try
      TMaster.FExamPanel:= 'IMU';
      TMaster.FExamDate:= FormatDateTime( 'yyyymmdd', now );
      TMaster.FInstTime:= FormatDateTime('yyyymmddhhnnss', now);
      TMaster.FRcvTime := Str2ViewDTM(TMaster.FInstTime);
      TGlobal.ErrMsg:='';

      for i:=0 to mmTemp.Lines.Count -1 do begin
          sData:= mmTemp.Lines.Strings[i];
          if sData = '' then continue;

          sFrame:= Trim(TokenStr( sData,',',1));

          if sFrame = 'V' then begin    //'E'->응급, 'R' -> RackPos, //'K' -> WorkSheet mode
              //TMaster.FInstInfo:= TokenStr(sData,',',1);
              //TMaster.FInstInfo:= TMaster.FInstInfo + ',' + TokenStr(sData,',',2);
              //TMaster.FInstInfo:= TMaster.FInstInfo + ',' + TokenStr(sData,',',3);
          end
          else
          if sFrame = 'K' then begin
              TMaster.FBarCode:= Trim(TokenStr(sData,',',2));
              uTas_OrderProcess(TMaster);
              exit;
          end
          else
          if (sFrame = 'S') or (sFrame = 's') or (sFrame = 'Q') then begin      //'Q' QC?
              //결과, MASTER
              TMaster.FExamSeq:= '';
              TMaster.FBarCode:= Trim(TokenStr(sData,',',7));
              TMaster.FRack   := Trim(TokenStr(sData,',',5));
              TMaster.FPos    := Trim(TokenStr(sData,',',6));

              if TMaster.FBarCode = '' then
                  TMaster.FBarCode:= Trim(TokenStr(sData,',',8));

              GrdIdx:= GetBCDIndex(TMaster);
              if TMaster.FExamSeq = '' then
                  TMaster.FExamSeq := DM.GetExamSeq( TMaster.FExamDate );

              DM.DownLoadOrder_SCHUH(TMaster);

              DM.SaveMaster(TMaster);
              DispOnePat_Idx(TMaster, GrdIdx);

              //RESULT
              ResCnt:= StrToIntDef(TokenStr(sData,',',10),0);

              for j:=0 to ResCnt-1 do begin
                  TMaster.FUpCode:= TokenStr(sData,',',10+(j*6)+1);
                  TMaster.FAbbr:= TCode.GetAbbr_Up(TMaster.FUpCode);
                  TMaster.FExamCode:= TCode.GetExamCode_UpCode(TMaster.FUpCode);

                  dRES:= StrToFloatDef(Trim(TokenStr(sData,',',10+(j*6)+2)),-1);
                  if dRES > 0 then
                      TMaster.FResult:= FloatToStr(dRES)
                  else
                      TMaster.FResult:= Trim(TokenStr(sData,',',10+(j*6)+2));

                  TMaster.FFlag  := TokenStr(sData,',',10+(j*6)+3);

                  if TMaster.FResult = 'NC' then begin
                      //L1
                      TMaster.FResult:= '<0.5';   //Flag:09
                      //L3
                  end;

                  if TMaster.FFlag = '00' then
                      TMaster.FFlag:= 'N';

                  if TMaster.FUpCode = '01' then begin
                      dAFP:= StrToFloatDef(TMaster.FResult, -1);
                      sAfp:= TMaster.FResult;
                  end
                  else
                  if TMaster.FUpCode = '02' then begin
                      dAfpL3:= StrToFloatDef(TMaster.FResult, -1);
                      sAfpL3:= TMaster.FResult;
                  end;

                  DM.SaveResult(TMaster);
                  DispOneResult(TMaster, GrdIdx);
                  //UploadOneExam(TMaster, GrdIdx);
              end;

              //계산식
              if sAfpL3 = '<0.5' then begin
                  TMaster.FUpCode:= '99';
                  TMaster.FAbbr:= TCode.GetAbbr_Up(TMaster.FUpCode);
                  TMaster.FExamCode:= TCode.GetExamCode_UpCode(TMaster.FUpCode);
                  TMaster.FResult:= '<0.3';
                  TMaster.FFlag  := 'L';
                  //UploadOneExam(TMaster, GrdIdx);
                  DM.SaveResult(TMaster);
                  DispOneResult(TMaster, GrdIdx);
              end
              else
              if (dAFP > 0) and (dAfpL3 > 0 ) then begin
                  TMaster.FUpCode:= '99';
                  TMaster.FAbbr:= TCode.GetAbbr_Up(TMaster.FUpCode);
                  TMaster.FExamCode:= TCode.GetExamCode_UpCode(TMaster.FUpCode);
                  TMaster.FResult:= Trim( Format('%5.1f', [ (dAFP * dAfpL3) / 100 ]) );
                  TMaster.FFlag  := 'N';
                  //UploadOneExam(TMaster, GrdIdx);
                  DM.SaveResult(TMaster);
                  DispOneResult(TMaster, GrdIdx);
              end;

              UploadProcess(TMaster, GrdIdx, False);

              gdIf.Refresh;
          end;
      end;



  finally
      TMaster.Free;
  end;
end;

procedure TF_Main.DispOnePat_ORD(TMaster:TIfMaster; GrdIdx:integer=0);
var
  VC, i, nCol, nRow:integer;
  Abr:string;
begin
  nRow:= AddOrdIndex;

  gdIf.AddCheckBox(0, nRow, False, False);
  gdIf.Cells[C_DTM, nRow] := FormatDateTime('yyyy-mm-dd hh:nn:ss', now);
  gdIf.Cells[C_SEQ, nRow] := '';
  gdIf.Cells[C_EDT, nRow] := '';
  gdIf.Cells[C_BCD, nRow] := TMaster.FBarCode;
  gdIf.Cells[C_POS, nRow] := TMaster.FRack+'-'+TMaster.FPos;
  gdIf.Cells[C_PID, nRow] := TMaster.FPID;
  gdIf.Cells[C_PNM, nRow] := TMaster.FPNM;
  //gdIf.Cells[C_PNL, nRow] := TMaster.FExamPanel;

  if TMaster.FOrdState = 'N' then
      gdIf.Cells[C_STA, nRow] := '오더없음'
  else
  if TMaster.FOrdState = 'Y' then
      gdIf.Cells[C_STA, nRow] := '오더전송';

  gdIf.AutoNumberCol(0);

  VC:= VarArrayDimCount(TMaster.vAbbr);
  if VC > 0 then begin
      for i:=0 to VarArrayHighBound(TMaster.vAbbr,1) do begin
         Abr:= Trim(TMaster.vAbbr[i]);

         if Abr <> '' then begin
             nCol:= GetAbbrIndex(gdIf, Abr);
             if (nCol > 0) and (gdIf.Cells[nCol, nRow] = '') then
                 gdIf.Cells[nCol, nRow]:= 'Y';
         end;
      end;
  end;
end;

function TF_Main.AddOrdIndex: integer;
begin
  Result:= 0;
  if (gdIf.RowCount = 2) and (gdIf.Cells[C_BCD,1]='') then begin
      Result:=1;  exit;
  end
  else begin
      gdIf.AddRow;
      Result:= gdIf.RowCount -1;
  end;

end;

procedure TF_Main.doG7(const cData: string);
var
  i:integer;
  cStr,cBuffer:string;
  sTmp:string;
  iNm:string;
begin
  iNm:= TGlobal.FIName;

   // 전역변수로 보관한 자료를 읽음
   cStr:='';
   cBuffer:=cData;
   for i:= 1 to Length(cBuffer) do begin
     case cBuffer[i] of
     ETX:begin
             vcTxData:= vcTxData+ETX;
             SendStr(ACK);
         end;
     EOT:begin
             Inc(gACK);
             sTmp:= vcTxData+'[EOT]';
             AddDebugLine(iNm+'->'+sTmp);
             SetCommData(vcTxData);
             G7_Rcv;
             sTmp:='';
             vcTxData:='';

             {if gACK = 2 then begin
                 try
                     SetCommData(vcTxData);
                     G7_Rcv;
                 finally
                     gACK:=0;
                     sTmp:='';
                     vcTxData:='';
                 end;
             end
             else
             if gACK = 3 then begin
                gACK:=0;
                sTmp:='';
                vcTxData:='';
             end
             else begin
                sTmp:='';
                vcTxData:='';
             end; }
         end;
         else begin
            vcTxData:=vcTxData+cBuffer[i];
         end;
     end;
   end;
end;

procedure TF_Main.G7_Rcv;
var
  TMaster:TIfMaster;
  i:integer;
  dVal:double;
  bUpSucc:boolean;
  UpSucc:integer;
  GrdIdx:integer;
  UpCd,Flg,Res:string;
  sTemp, ErrMsg:string;
begin
  ErrMsg:= '';

  TMaster:= TIfMaster.Create;

  try // ~Finally
      sTemp:= mmTemp.Text; //Trim( mmTemp.Text );
      if (Length(sTemp) < 80) or (Length(sTemp) > 90) then exit;

      TMaster.FExamDate:= FormatDateTime( 'yyyymmdd', now );
      TMaster.FInstTime:= FormatDateTime('yyyymmddhhnnss', now);
      TMaster.FRcvTime := Str2ViewDTM(TMaster.FInstTime);

      TMaster.FBarCode:= Trim(Copy(sTemp, 65, 20));
      TMaster.FPos    := Copy(sTemp, 5, 8);
      if TMaster.FBarCode = '' then
          TMaster.FBarCode:= TMaster.FPos;

      GrdIdx:= GetAddIndex(gdIf);
      TMaster.FExamSeq := DM.GetExamSeq( TMaster.FExamDate );

      TMaster.FOrdState:= 'N';
      if DM.DownLoadPAT_JAIN(TMaster) then
          if DM.DownLoadOrder_JAIN(TMaster) then
              TMaster.FOrdState:= 'Y';

      DM.SaveMaster(TMaster);

      DispOnePat_Idx(TMaster, GrdIdx);

      TMaster.FFlag  := Copy(sTemp, 63, 2);
      if TMaster.FFlag = '00' then
          TMaster.FFlag:= 'N';

      if TMaster.FFlag <> 'N' then
          TMaster.FUpState:= 'X';

      for i:=0 to 9 do begin
          TMaster.FIfCode:= IntToStr(i+1);
          TMaster.FResult:= Trim(Copy(sTemp, (i*5)+1+13, 5));
          UploadOneExam(TMaster, GrdIdx);
      end;

      //Flag를 결과처럼 보여주자..
      TMaster.FIfCode:= IntToStr(11);
      TMaster.FResult:= TMaster.FFlag;
      TMaster.FFlag  := TMaster.FFlag;
      UploadOneExam(TMaster, GrdIdx);

      TGlobal.ErrMsg:='';

      UploadProcess(TMaster, GrdIdx, True);

      gdIf.Refresh;
  finally
      TMaster.Free;
  end;

end;

procedure TF_Main.UploadProcess_One(TObj: TObject; GrdIdx:integer);
var
  TMaster:TIfMaster absolute TObj;
begin
  if mnAuto.Checked = True then begin
      if DM.UploadResult_Direct(TMaster) then
          TMaster.FUpState:= 'Y'
      else
          TMaster.FUpState:= 'N';
  end;

  DM.SaveState( TMaster.FExamDate, TMaster.FExamSeq, '', TMaster.FUpState);
  F_Main.DispOnePat_ALL(TMaster, GrdIdx);

end;

procedure TF_Main.DoTest1(sData: string);
var
 sTemp, sBuffer:string;
 i: Integer;
begin
   sTemp:= sData;

   //<STX>Q01123456789<ETX>

   for i:=1 to Length(sTemp) do begin
       Case sTemp[i] of
           STX: begin
                    vcTxData:='';
                    sBuffer :='';
                end;
           ETX: begin
                    mmTemp.Text:= vcTxData + sBuffer;
                    Test1_Process(mmTemp.Text);
                    vcTxData:='';
                    sBuffer:='';
                    exit;
                end;
           else begin
                    sBuffer:= sBuffer + sTemp[i];
                end;
       end;
   end;

   vcTxData:= vcTxData + sBuffer;

end;

procedure TF_Main.Test1_OrderProcess(Barcode: string);
begin
  exit;
end;

procedure TF_Main.Test1_Process(sData: string);
var
  Flag, BCD:string;
begin
  TGlobal.DataLog:= sData;

  if Trim(sData) = '' then exit;

  Flag:= Copy(sData,1,1);
  BCD := Trim(Copy(sData,4,15));

  Case Flag[1] of
      'Q': TEST1_OrderProcess(BCD);       //<STX>Q01123456789
      'R': TEST1_ResultProcess(sData);    //<STX>R01123456789 0103050012f <ETX>
  end;


end;

procedure TF_Main.Test1_ResultProcess(sData: string);
var
  TMaster: TIfMaster;
  ResStr, sTmp:string;
  i, GrdIdx:integer;
  dRes:double;
  bUpSucc:boolean;
begin
  if sData = '' then exit;

  TMaster:= TIfMaster.Create;
  try
      TMaster.FExamDate:= FormatDateTime( 'yyyymmdd', now );
      TMaster.FExamSeq := DM.GetExamSeq(TMaster.FExamDate);
      TMaster.FInstTime:= FormatDateTime('yyyymmddhhnnss', now);
      TMaster.FRcvTime := Str2ViewDTM(TMaster.FInstTime);

      grdIdx:= GetAddIndex(gdIf);

      TMaster.FRack:= Trim(Copy(sData,19,2));
      TMaster.FPos := Trim(Copy(sData,21,2));

      sTmp:= ReplaceAllStr(Copy(sData,1,20),' ','');
      TMaster.FBarCode:= Trim(Copy(sData,4,11));
      if TMaster.FBarCode = '' then
          TMaster.FBarCode:= TMaster.FPOS;

      if Copy(TMaster.FBarCode,1,1) = 'C' then
          TMaster.FQCYN:= 'Y';

      ResStr:= Copy(sData,25,4);
      ResStr:= ReplaceAllStr(ResStr,'*','');

      if (ResStr = '-004') or (ResStr ='0-04') then
          TMaster.FResult:= 'NF'
      else
      if (ResStr = '-002') or (ResStr='0-02') then
          TMaster.FResult:= 'NR'
      else begin
          dRes:= StrToFloatDef(ResStr,-1);
          if dRes >= 0 then
              TMaster.FResult:= FloatToStr(dRes);
      end;

      TMaster.FIfCode:= 'ESR';
      TMaster.FAbbr:= 'ESR';
      TMaster.FExamPanel:= 'ESR';

      if Pos('*', TMaster.FResult)> 0 then
          TMaster.FFlag:= '*'
      else
      if (TMaster.FResult='NF') or (TMaster.FResult='NR') then
          TMaster.FFlag:='H'
      else
          TMaster.FFlag := 'N';

      TGlobal.ErrMsg:='';

      if Pos('LATEX', TMaster.FBarCode) > 0 then
          TMaster.FQCYN:= 'Y';

      DM.DownLoadOrder_JEIL(TMaster);

      DM.SaveMaster(TMaster);
      DM.SaveResult(TMaster);

      if (TMaster.FResult = 'NR') or (TMaster.FResult = 'NF') then
      begin
          DispOnePat_ALL(TMaster, GrdIdx);
      end
      else
          UploadProcess_One(TMaster, GrdIdx);

      gdIf.Refresh;

  finally
      TMaster.Free;
  end;
end;

procedure TF_Main.Test1_SendNoOrder(Barcode: string);
begin

end;

function TF_Main.Test1CheckSum(sData: string): string;
var
  ix1  : integer;
  sXor : string;
begin
  sXor := '';
  sXor := copy(sData, 1, 1);
  for ix1 := 2 to Length(sData) do
  begin
    sXor := chr( Ord(sXor[1]) Xor Ord(copy(sData, ix1, 1)[1]) )
  end;

  //만약에 마지막이 ETX라면 Hex-F7로 변환한다. 16진수(F7) -> 10진수(127)

  if sXor = chr(3) then
    sXor := chr(127);

  Result := sXor;


end;

procedure TF_Main.DispOnePat_ALL(TMaster: TIfMaster; GrdIdx:integer);
var
  ExamTime:string;
  nRow:integer;
begin
  if (GrdIdx >= gdIf.RowCount) or (grdIdx<1) then exit;

  nRow:= GrdIdx;

  gdIf.AddCheckBox(0, nRow, False, False);
  gdIf.Cells[C_SEQ, nRow] := TMaster.FExamSeq;
  gdIf.Cells[C_EDT, nRow] := TMaster.FExamDate;
  gdIf.Cells[C_DTM, nRow] := TMaster.FRcvTime;
  gdIf.Cells[C_BCD, nRow] := TMaster.FBarCode;
  //if TMaster.FAbbr <> '' then
  //    gdIf.Cells[C_ABR, nRow] := TMaster.FAbbr
  //else
  //    gdIf.Cells[C_ABR, nRow] := TMaster.FIfCode;

  gdIf.Cells[C_PID, nRow] := TMaster.FPID;
  gdIf.Cells[C_PNM, nRow] := TMaster.FPNM;
  //gdIf.Cells[C_RES, nRow] := TMaster.FResult;

  //if TMaster.FFlag <> 'N' then
  //    gdIf.FontColors[C_RES, nRow]:= clRed;

  if TMaster.FRack <> '' then
      gdIf.Cells[C_POS, nRow] := TMaster.FRack+'-'+ TMaster.FPos
  else
      gdIf.Cells[C_POS, nRow] := TMaster.FPos;

  if TMaster.FUpState = 'Y' then
      gdIf.Cells[C_STA, nRow] := '전송완료'
  else begin
      gdIf.SetCheckBoxState(0, nRow, True);

      if TMaster.FOrdState = 'N' then
          gdIf.Cells[C_STA, nRow] := '오더없음'
      else
      if TMaster.FUpState = 'X' then begin
          gdIf.Cells[C_STA, nRow] := '결과이상';
          gdIf.SetCheckBoxState(0, nRow, False);
      end
      else
          gdIf.Cells[C_STA, nRow] := '미전송';
  end;
  gdIf.AutoNumberCol(0);

end;

procedure TF_Main.N10Click(Sender: TObject);
begin
    WinExec('"C:\Program Files\Internet Explorer\iexplore.exe" "http://sansoft.anyhelp.net"', CmdShow);
end;

function TF_Main.GetChorusRemark(IFCD, RES, FLG: string): string;
begin
  if IFCD = 'MYCO-M' then begin
      //Result
      if FLG = 'N' then begin
          Result:= '';
      end
      else
      if FLG = 'P' then begin
          Result:= '* IgM '+RES+'   (참고치) Positive > 1.1';
      end
      else
      if FLG = 'B' then begin
          Result:= '* IgM '+RES+'   (참고치) Borderline 0.9 - 1.1';
      end;
  end;

end;

function TF_Main.GetChorusResult(IFCD, RES, FLG: string): string;
begin
  Result:= RES +' / '+ FLG;
end;

procedure TF_Main.N11Click(Sender: TObject);
var
  OID, NId:string;
begin
  OID:= TGlobal.FUserID;
  NID:= InputBox('검사자 설정', '변경하실 사번을 입력하세요', OID);
  TGlobal.FUserID:= NID;
  TGlobal.SaveIni;
  MyFormCaptionChange(mnAuto.Checked);
end;

procedure TF_Main.gdIfGetCellColor(Sender: TObject; ARow, ACol: Integer;
  AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
var
  STA:string;
begin
  if (ARow > 0) and (ACol = C_STA) then begin
      STA:= gdIf.Cells[ACol, ARow];
      if STA <> '전송완료' then
          AFont.Color:= clRed
      else
          AFont.Color:= clBlack;
  end;
end;

procedure TF_Main.Chorus_Trio_ResProcess(sData:string);
var
  TMaster:TIfMaster;
  dVal:double;
  bUpSucc:boolean;
  UpSucc:integer;
  GrdIdx:integer;
  UpCd,Flg,Res:string;
  sTemp, ErrMsg:string;
begin
  ErrMsg:= '';

  TMaster:= TIfMaster.Create;

  try // ~Finally
      sTemp:= sData; //Trim( mmTemp.Text );

      if Length(sTemp) < 13 then exit;

      TMaster.FExamDate:= FormatDateTime( 'yyyymmdd', now );
      TMaster.FExamSeq := DM.GetExamSeq( TMaster.FExamDate );
      TMaster.FExamPanel:= 'IMU';
      TGlobal.ErrMsg:='';

      TMaster.FInstTime:= FormatDateTime('yyyymmddhhnnss', now);
      TMaster.FRcvTime := Str2ViewDTM(TMaster.FInstTime);

      TMaster.FBarCode:= Trim(Copy(sTemp, 4, 18));
      UpCd := Trim(Copy(sTemp,23,7));
      Flg  := Trim(Copy(sTemp,30,1));
      Res  := Trim(Copy(sTemp,31,12));

      TMaster.FIfCode:= UpCd;
      TMaster.FAbbr  := TCode.GetAbbr_IF(UpCd);
      TMaster.FFlag  := Flg;
      TMaster.FRMK   := '';
      TMaster.FResult:= GetChorusResult(UpCd, Res, Flg);

      //QC추가
      if Length(TMaster.FBarCode) = 18 then begin
          TMaster.FResult:= Res;
          TMaster.FQCYN:= 'Y';
          SaveChorus_QC(TMaster);
          //exit;
      end;
      ///

      GrdIdx:= GetAddIndex(gdIf);

      if DM.DownLoadOrder_SCHUH(TMaster) = False then
          TMaster.FExamCode:= TCode.GetExamCode_UpCode(TMaster.FUpCode);

      DM.SaveMaster(TMaster);

      DM.SaveResult(TMaster);
      UploadProcess_One(TMaster, GrdIdx);

      gdIf.Refresh;
  finally
      TMaster.Free;
  end;

end;

procedure TF_Main.Alegria_Rcv;
var
  TMaster:TIfMaster;
  i, GrdIdx:integer;
  sData,cFrame,sTemp,sInstExamTime, ResSeq,
  sSampleID,sPanel:string;
  bUpSucc:boolean;
  SampleSeq, SampleCount, UpSucc:integer;
  Res_Val, Res_Txt, Res_Unt:string;
begin
  SampleSeq:=0;
  SampleCount:=0;
  GrdIdx:=0;

  TMaster:= TIfMaster.Create;

  try // ~Finally
      for i:=0 to mmTemp.Lines.Count -1 do begin
          sData:= Trim( TokenStr( mmTemp.Lines.Strings[i],CR,1 ) );
          if Trim( sData ) = '' then
              continue;

          cFrame:=Copy( TokenStr( sData,'|' ,1 ), 2, 1 );

          if cFrame='H' then begin   //1H|\^&|||BIOSITE00045480|||||||P|LIS6|20080530162200|
              Inc(SampleCount);

              TMaster.FExamDate:= FormatDateTime( 'yyyymmdd', now );
              TMaster.FExamSeq := DM.GetExamSeq( TMaster.FExamDate );
              GrdIdx:= GetAddIndex(gdIf);
          end
          else if cFrame='P' then begin
              sData:= sData+'|';
              SampleSeq:= StrToIntDef(Trim(TOkenStr(sData,'|',2)), 1);  //사용안함.나중을위해서.
              sSampleID:=  Trim(TokenStr(sData,'|',4));
              if Pos('QC', sSampleID) > 0 then exit;
          end
          else if cFrame = 'O' then begin
              bUpSucc:=False;
          end
          //결과
          else if cFrame='R' then begin
              //2R|1|^^^SS-B|N/A^N/A^QC Failed|||||F||TEST||20121113151326

              TMaster.FIfCode  := Trim(TokenStr(TokenStr(sData,'|', 3),'^', 4));
              TMaster.FAbbr    := TCode.GetAbbr_IF(TMaster.FIfCode);

              Res_Val:= Trim(TokenStr(TokenStr(sData,'|', 4),'^',1));
              Res_Unt:= Trim(TokenStr(TokenStr(sData,'|', 4),'^',2));
              Res_Txt:= Trim(TokenStr(TokenStr(sData,'|', 4),'^',3));

              if Res_Val = 'N/A' then begin
                  TMaster.FFlag:= 'X';
                  TMaster.FResult:= Res_Txt;
              end
              else begin
                  TMaster.FFlag:= Copy(Res_Txt,1,1);
                  TMaster.FResult:= GetAlegriaResText(TMaster.FIfCode, Res_Val, Res_Txt);
              end;

              ResSeq:= TokenStr(sData,'|',2);
              if ResSeq = '1' then begin
                  TMaster.FInstTime:= TokenStr(sData,'|',13);
                  TMaster.FRcvTime := Str2ViewDTM(TokenStr(sData,'|',13));
                  TMaster.FBarCode  := sSampleId;
                  DM.DownLoadOrder_JND_RES(TMaster);
                  DM.SaveMaster(TMaster);
              end;

              DM.SaveResult(TMaster);
              UploadProcess_One(TMaster, GrdIdx);
          end

          //Last
          else if cFrame='L' then begin
               TMaster.Clear;
          end;
      end;
  finally
      TMaster.Free;
  end;
end;

procedure TF_Main.DoAlegria(const cData: string);
var
  i:integer;
  cStr,cBuffer:string;
  INM, sTmp:string;
begin
   // 전역변수로 보관한 자료를 읽음
   INM:= TGlobal.FIName;

   cStr:='';
   cBuffer:=cData;

   for i:= 1 to Length(cBuffer) do begin
     case cBuffer[i] of
     ACK:begin
             AddDebugLine(INM+'->[ACK]');
             SendStr(EOT);
         end;
     ENQ:begin
             AddDebugLine(INM+'->[ENQ]');
             SendStr(ACK); //ENQ
             vcTxData:='';
             sTmp:='';
         end;
     STX:begin
             ;
         end;
     ETB:begin
             vcTxData:= vcTxData+ETB;
             SendStr(ACK);
         end;

     ETX:begin
             vcTxData:= vcTxData+ETX;
             SendStr(ACK);
         end;
     EOT:begin
             AddDebugLine(INM+'->'+sTmp);
             try
                 SetCommData(vcTxData);
                 Alegria_Rcv;
             finally
                 sTmp:='';
                 vcTxData:='';
             end;
         end;
         else begin
            vcTxData:=vcTxData+cBuffer[i];
         end;
     end;
   end;

end;

function TF_Main.GetAlegriaResText(IFCD, RES, FLG: string): string;
var
  Txt:string;
begin
  Txt:= Copy(FLG,1,1);
  if Txt = 'P' then
      Result:= 'Positive('+RES+')'
  else
  if Txt = 'N' then
      Result:= 'Negative('+RES+')'
  else
  if Txt = 'B' then
      Result:= 'Borderline('+RES+')'
  else
      Result:= FLG+'('+RES+')';

end;

procedure TF_Main.InitGlobal;
begin
  FillChar(gTxbuffer, SizeOf(gTxbuffer), 0);

  gInputCount:= 0;
  gNeedCount:= 0;
  gCMDCode:= 0;

end;

procedure TF_Main.SendChorus_ACK;
begin
  SendStr('');
end;

function TF_Main.ByteArrayToStr(aRcv: array of Byte): string;
var
  i:integer;
  iLow, iHigh:integer;
  S:string;
begin
  Result:='';

  iLow:= Low(aRcv);
  iHigh:= High(aRcv);

  //델파이 스트링에서 #0은 처리하지 않으므로.. Space로 변경.
  for i:=iLow to iHigh do begin
    if (aRcv[i] = 0) or (aRcv[i]>127) then
        aRcv[i]:= $20
  end;

  for i:=iLow to iHigh do begin
      S:= S + Char(aRcv[i]);
  end;

  Result:= S;

end;

procedure TF_Main.Chorus_Trio_Rcv;
var
  i:integer;
  sTemp:string;
begin
  Screen.Cursor:= crHourGlass;
  try
      for i:=0 to slRcv.Count -1 do begin
          sTemp:= slRcv.Strings[i]; //Trim( mmTemp.Text );

          if Length(sTemp) < 13 then Continue;

          Chorus_Trio_ResProcess(sTemp);
      end;
  finally
      Screen.Cursor:= crDefault;
  end;

end;

procedure TF_Main.SaveChorus_QC(TMaster: TIfMaster);
var
  dMin, dMax, m, sd:double;
begin
  //QC추가
  if Copy(TMaster.FBarCode,17,2) = '01' then
      TMaster.FTyp:= 'Q'
  else
      TMaster.FTyp:= 'C';

  TMaster.FLotNo:= Copy(TMaster.FBarCode,4,3);
  TMaster.FLotMin:= Copy(TMaster.FBarCode,8,2);
  TMaster.FLotMax:= Copy(TMaster.FBarCode,10,2);

  if TMaster.FTyp = 'Q' then
      TMaster.FLotLev:= '1'
  else
      TMaster.FLotLev:= '4';   //CAL

  //Mean, SD값 구한다. 수치형 결과를 가지고 있어야 한다.
  DM.SetExamLotMean(TMaster);

  if DM.CheckLot('CHORUS', TMaster.FIfCode, TMaster.FLotNo, TMaster.FTyp, TMaster.FLotLev) = False then
      DM.SaveOneLotInfo('CHORUS'
                       , TMaster.FLotNo
                       , TMaster.FLotLev
                       , TMaster.FIfCode
                       , TMaster.FLotMean
                       , TMaster.FLotSD
                       , ''
                       , ''
                       , TMaster.FLotMin
                       , TMaster.FLotMax
                       , TMaster.FTyp);

  //ExamDate, ExamSEQ를 새로 받는걸로 처리하는것이 좋을듯..
  DM.SaveQC(TMaster);
  //exit;

  //NegativeQC
  if TMaster.FTyp = 'Q' then begin
     TMaster.FLotNo:= TMaster.FLotNo;
     TMaster.FLotMin:= '0.0';
     TMaster.FLotMax:= '0.1';
     TMaster.FLotLev:= '2';

     if (StrToInt(FormatDateTime('ss', now)) mod 2) = 0 then
         TMaster.FResult:= '0.0'
     else
         TMaster.FResult:= '0.1';

     dMin:= StrToFloat(TMaster.FLotMin);
     dMax:= StrToFloat(TMaster.FLotMax);
     TMaster.FLotMin:= Trim(Format('%5.1f', [dMin]));
     TMaster.FLotMax:= Trim(Format('%5.1f', [dMax]));

     m:= mean([dMin, dMax]);
     sd:= stddev([dMin, dMax]);

     TMaster.FLotMean:= Trim(Format('%5.1f', [m]));
     TMaster.FLotSD:= Trim(Format('%5.1f', [sd]));


      if DM.CheckLot('CHORUS', TMaster.FIfCode, TMaster.FLotNo, TMaster.FTyp, TMaster.FLotLev) = False then
          DM.SaveOneLotInfo('CHORUS'
                           , TMaster.FLotNo
                           , TMaster.FLotLev
                           , TMaster.FIfCode
                           , TMaster.FLotMean
                           , TMaster.FLotSD
                           , ''
                           , ''
                           , TMaster.FLotMin
                           , TMaster.FLotMax
                           , TMaster.FTyp);

      DM.SaveQC(TMaster);
  end;
end;

procedure TF_Main.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  TGlobal.MainTop:= Self.Top;
  TGlobal.MainLeft:= Self.Left;
  TGlobal.MainHeight:= Self.Height;
  TGlobal.MainWidth:= Self.Width;
end;

procedure TF_Main.DB1Click(Sender: TObject);
begin
  if F_Server = nil then
      F_Server:= TF_Server.Create(nil);

  F_Server.Show;
end;

procedure TF_Main.gdIfClickCell(Sender: TObject; ARow, ACol: Integer);
var
  i:integer;
  b:boolean;
begin
  if (ARow=0) and (ACol=0) then begin
      for i:=1 to gdIf.RowCount-1 do begin
          if i=1 then begin
              b:= False;
              gdIf.GetCheckBoxState(0, i, b);
          end;

          gdIf.SetCheckBoxState(0, i, Not b);
      end;
  end
end;

procedure TF_Main.N21Click(Sender: TObject);
begin
    WinExec('C:\Program Files\Internet Explorer\iexplore.exe" "http://downloadus1.teamviewer.com/download/TeamViewerQS.exe"', CmdShow);
end;

procedure TF_Main.doEPOC_BT(const cData: string);
var
  i:integer;
  cStr,cBuffer:string;
  sTmp:string;
  iNm:string;
begin
  iNm:= TGlobal.FIName;

   // 전역변수로 보관한 자료를 읽음
   cStr:='';
   cBuffer:=cData;
   for i:= 1 to Length(cBuffer) do begin
     case cBuffer[i] of
         STX:begin
                 vcTxData:= '';
             end;
         //CR:begin
         //       vcTxData:= vcTxData +
         //   end;
         GS:begin
                SetCommData(vcTxData);
                EPOC_BT_Rcv;
                vcTxData:='';
                exit;
             end;
         else begin
                vcTxData:=vcTxData+cBuffer[i];
         end;
     end;
   end;

end;

procedure TF_Main.EPOC_BT_Rcv;
var
  TMaster:TIfMaster;
  sData, SmpId:string;
  i, GrdIdx:integer;
  dVal:double;
begin

  TMaster:= TIfMaster.Create;
  GrdIdx:=0;

  try // ~Finally
      TMaster.FExamPanel:= 'EPOC';
      TMaster.FExamDate:= FormatDateTime( 'yyyymmdd', now );
      TMaster.FInstTime:= FormatDateTime('yyyymmddhhnnss', now);
      TMaster.FRcvTime := Str2ViewDTM(TMaster.FInstTime);


      for i:=0 to mmTemp.Lines.Count -1 do begin
          {
          pH       7.089         Low
          pCO2     50.6   mmHg   High
          pO2      88.9   mmHg
          Na+      113    mmol/L Low
          K+       2.1    mmol/L Low
          Ca++     1.50   mmol/L High
          Glu      39     mg/dL  Low
          Lac      0.95   mmol/L
          Hct      < 10   %      Low
          }
          sData:= Trim(mmTemp.Lines.Strings[i]);
          if sData = '' then continue;

          //if sData = 'epoc BGEM Blood Test'
          if (sData = 'No Results: Corrected') or (sData='Reference Ranges') then
              Break;

          if Pos('Patient ID:', sData) > 0 then begin
              SmpID:= UpperCase(Trim(Copy(sData,12,20)));
              TMaster.FBarCode:= F_Work.GetWorkBcd;
              if TMaster.FBarCode = '' then
                  TMaster.FBarCode:= SmpID;

              TMaster.FPos:= SmpID;
              TMaster.FExamSeq:= DM.GetExamSeq(TMaster.FExamDate);

              DM.DownLoadOrder_DJI(TMaster);

              GrdIdx:= GetAddIndex(gdIf);

              DispOnePat_Idx(TMaster, GrdIdx);
              DM.SaveMaster(TMaster);
          end;

          TMaster.FUpCode:= Trim(Copy(sData,1,9));

          if TCode.SetCode_UpCode(TMaster.FUpCode) then begin
              TMaster.FResult:= Trim(Copy(sData, 10, 7));
              {
              dVal:= StrToFloatDef(TMaster.FResult, -100);
              if dVal > -99 then begin
                  if TMaster.FUpCode = 'pH' then
                      TMaster.FResult:= Trim(Format('%5.2f', [dVal]))
                  else if TMaster.FUpCode = 'pCO2' then
                      TMaster.FResult:= Trim(Format('%5.0f', [dVal]))
                  else if TMaster.FUpCode = 'pO2' then
                      TMaster.FResult:= Trim(Format('%5.0f', [dVal]))
                  else if TMaster.FUpCode = 'cHCO3-' then
                      TMaster.FResult:= Trim(Format('%5.0f', [dVal]))


              end;}

              TMaster.FAbbr:= TCode.GetAbbr_Up(TMaster.FUpCode);
              TMaster.FExamCode:= TCode.GetExamCode_UpCode(TMaster.FUpCode);

              //Unit:= 16,7
              TMaster.FFlag:= Trim(Copy(sData,23,20));
              if TMaster.FFlag = 'Low' then
                  TMaster.FFlag:= 'L'
              else
              if TMaster.FFlag = 'High' then
                  TMaster.FFlag:= 'H'
              else
              if TMaster.FFlag = '' then
                  TMaster.FFlag:= 'N';

              DM.SaveResult(TMaster);
              DispOneResult(TMaster, GrdIdx);

          end;
      end;

      UploadProcess(TMaster, GrdIdx, True);
      gdIf.Refresh;

  finally
      TMaster.Free;
  end;

end;

procedure TF_Main.btnWorkClick(Sender: TObject);
begin
  F_Work.Show;
end;

Initialization
  TGlobal:= TGlobalVar.Create;
  TConnection:= TDbConnection.Create;
  TCode:= TCodeInfo.Create;
  slrcv:= TStringList.Create;
  slSender:= TStringList.Create;

  gACK:=0;

finalization
  slSender.Free;
  slrcv.Free;

  TCode.Free;
  TConnection.Free;

  TGlobal.SaveIni;
  TGlobal.Free;
end.
