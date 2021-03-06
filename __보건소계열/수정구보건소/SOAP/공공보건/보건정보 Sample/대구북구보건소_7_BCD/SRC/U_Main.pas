{
  ?뱸 ?Ϻ??Ǽ? ?˻???
  ?????? ?????? 053-665-3241
}
unit U_Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Buttons, Menus, OleCtrls, ComCtrls, ExtCtrls,
  U_IfClass, StdCtrls, Grids, BaseGrid, AdvGrid, CPort, CPortCtl;

type
  TF_Main = class(TForm)
    Panel2: TPanel;
    StatusBar1: TStatusBar;
    MainMenu1: TMainMenu;
    pnLog: TPanel;
    mmTemp: TMemo;
    gdIf: TAdvStringGrid;
    N1: TMenuItem;
    N1_1: TMenuItem;
    N1_4: TMenuItem;
    L1: TMenuItem;
    DEBUG1: TMenuItem;
    mmLog: TMemo;
    N2: TMenuItem;
    N3: TMenuItem;
    Panel1: TPanel;
    Panel3: TPanel;
    Panel5: TPanel;
    gdResult: TAdvStringGrid;
    GroupBox1: TGroupBox;
    mmView: TMemo;
    N4: TMenuItem;
    btnTest: TButton;
    N6: TMenuItem;
    CLEAR1: TMenuItem;
    ComPort1: TComPort;
    pnSvr: TPanel;
    N7: TMenuItem;
    Rcv1: TMenuItem;
    pnPort: TPanel;
    pnBcd: TPanel;
    Panel6: TPanel;
    edOld: TEdit;
    edNew: TEdit;
    btnBcdChange: TButton;
    btnBcdClose: TButton;
    Panel7: TPanel;
    Panel8: TPanel;
    lbRow: TLabel;
    procedure btnTestClick(Sender: TObject);
    procedure DEBUG1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure gdIfGetAlignment(Sender: TObject; ARow, ACol: Integer;
      var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure gdIfGetCellColor(Sender: TObject; ARow, ACol: Integer;
      AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
    procedure ComPort1RxChar(Sender: TObject; Count: Integer);
    procedure N1_4Click(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure N3Click(Sender: TObject);
    procedure gdIfClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure gdIfCanEditCell(Sender: TObject; ARow, ACol: Integer;
      var CanEdit: Boolean);
    procedure N6Click(Sender: TObject);
    procedure mmViewDblClick(Sender: TObject);
    procedure gdResultGetCellColor(Sender: TObject; ARow, ACol: Integer;
      AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
    procedure CLEAR1Click(Sender: TObject);
    procedure N1_1Click(Sender: TObject);
    procedure N7Click(Sender: TObject);
    procedure Rcv1Click(Sender: TObject);
    procedure gdIfDblClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure btnBcdCloseClick(Sender: TObject);
    procedure btnBcdChangeClick(Sender: TObject);
    procedure Panel6MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure pnBcdMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
  private
    { Private declarations }
    procedure GridSetting;
    procedure FormSizeSetting;
    procedure DoH7180(const cData:string);
    procedure AddDebugLine(cStr:string);
    procedure ClearCommData;
    procedure SetCommData(const cStr:string);
    procedure SendStr(cStr:string);
    procedure DispOnePat_Order(TMaster:TH7180If);
    procedure DispOnePat_Result(TMaster:TH7180If);
    procedure DispOrderList(TMaster:TH7180If);
    procedure DispResultList(TMaster:TH7180If);
    procedure DispOnePatInfo(TMaster:TH7180If);
    procedure ChangeUpState(BarCode, UpState:string);
    function GetAbbrIndex(Abbr:string):integer;
    procedure AddResultGrid(Abbr, Res:string);
    procedure ClearResultGrid;
  public
    { Public declarations }
    function H747Str(cData:string): string;
    function H747CheckSum(SData:string): string;
    procedure H7180_Rcv;
    function SendSPMFrame:string;
    procedure SendOrderFrame(DataStr:string);
    procedure H7180_ResultProcess(sData:string);
    function PortSetup:boolean;
    function LoadComPort:boolean;
    function GetGridIndex(var TGrid:TAdvStringGrid; ExamDate, ExamSeq:string):integer;  overload;
    function GetGridIndex(TGrid:TAdvStringGrid; BarCode:string):integer; overload;
    function FindGridIndex(TGrid:TAdvStringGrid; BarCode:string):integer;
    procedure AddViewLog(Str:string);
  end;

type
  TMyThread = Class(TThread)
    constructor Create(CreateSuspended: Boolean);
  private
    procedure Run_ResultProcess;
  protected
    procedure Execute; override;
  end;

  EPortError = Class(EAbort)
    constructor Create;
  End;
var
  F_Main: TF_Main;
  ResultString:string;

implementation

uses GlobalVar, U_CodeInfo, SetDataBase, U_DM, StringLib, U_CommSet,
  U_TEST, U_CODE_SET;

const
  C_CHK  = 0;
//  C_DTM  = C_CHK+1;
//  C_SEQ  = C_DTM+1;
  C_BCD  = C_CHK+1;
  C_PID  = C_BCD+1;
  C_PNM  = C_PID+1;
  //C_ADT  = C_PNM+1;  //????????
  //C_ANO  = C_ADT+1;  //??????ȣ
  C_STA  = C_PNM+1;

  CommDelay = 100;

var
  vcTxData:string;

{$R *.dfm}

{ TForm1 }

function TF_Main.GetAbbrIndex(Abbr:string):integer;
var
  i:integer;
  nCol:integer;
begin
  Result:=0;

  for i:=1 to gdIf.AllColCount-1 do begin
      if Abbr = gdIf.Cells[i,0] then begin
          Result:= i;
          exit;
      end;
  end;

end;

procedure TF_Main.DispOnePat_Order(TMaster:TH7180If);
var
  nRow:integer;
begin
  nRow:= GetGridIndex(gdIf, TMaster.BarCode);

  if nRow > 0 then begin
      gdIf.AddCheckBox(0, nRow, False, False);
      with TMaster do begin
          //gdIf.Cells[C_SEQ, nRow] := FExamSeq;
          //gdIf.Cells[C_DTM, nRow] := ViewDateTime(FExamDate+FExamTime);
          gdIf.Cells[C_BCD, nRow] := BarCode;
          gdIf.Cells[C_PID, nRow] := FPatId;
          gdIf.Cells[C_PNM, nRow] := FPatNm;

          //gdIf.Cells[C_ADT, nRow] := FAcptDt;
          //gdIf.Cells[C_ANO, nRow] := FAcptNo;

          if TMaster.FOrdState = 'Y' then
              gdIf.Cells[C_STA, nRow] := '????????'
          else
              gdIf.Cells[C_STA, nRow] := '????????'
      end;
  end;

  //gdIf.OnClickCell(nil, nRow, C_DTM);
end;

procedure TF_Main.btnTestClick(Sender: TObject);
begin
  //DoH7180('>3E');
  H7180_Rcv;
end;

procedure TF_Main.DEBUG1Click(Sender: TObject);
begin
  DEBUG1.Checked:= DEBUG1.Checked;
  pnLog.Visible:= Debug1.Checked;
  btnTest.Visible:= DEBUG1.Checked;
end;

procedure TF_Main.FormShow(Sender: TObject);
begin
  GridSetting;
  FormSizeSetting;

  Self.Caption:= TGlobal.AppTitle;

  try
      LoadComPort;
  except
      ShowMessage('??Ʈ?????? Ȯ???ϼ???!');
      pnPort.Visible:= True;
  end;
  //else
  //    ComPort1.Open;
      //MsComm1.PortOpen:= True

  if Not TGlobal.HostConnecting then
      pnSvr.Visible:= True;
end;

procedure TF_Main.FormSizeSetting;
begin
  try
      Self.Top:= TGlobal.MainTop;
      Self.Left:= TGlobal.MainLeft;
      Self.Width:= TGlobal.MainWidth;
      Self.Height:= TGlobal.MainHeigh;
  except
      Self.Top:= DEFTOP;
      Self.Left:= DEFLFT;
      Self.Width:= DEFWID;
      Self.Height:= DEFHEI;
  end;

end;

function TF_Main.PortSetup: boolean;
begin
  Result:= False;
  pnPort.Visible:= False;

  {CPort}
  ComPort1.ShowSetupDialog;
  ComPort1.StoreSettings(stIniFile, TGlobal.AppPath+IniFileName);
  if ComPort1.Connected then
      Result:= True

  {MSComm}
  {if F_CommSet = nil then
      F_CommSet:= TF_CommSet.Create(Self);

  F_CommSet.TCommPort:= MSComm1;
  if F_CommSet.ShowModal = mrOk then
      Result:= True
  }

end;

procedure TF_Main.gdIfGetAlignment(Sender: TObject; ARow, ACol: Integer;
  var HAlign: TAlignment; var VAlign: TVAlignment);
begin
  VAlign:= vtaCenter;
  HAlign:= taCenter;
end;

procedure TF_Main.gdIfGetCellColor(Sender: TObject; ARow, ACol: Integer;
  AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
var
  LH,Abbr,sResult,sCode:string;
  DP:string;
  nCol:integer;
  sState, sNull:string;
begin
  if ARow = 0 then begin
  end
  else begin
      nCol:= gdIf.AllColCount;

      if ACol in [C_BCD..nCol -1] then begin
          AFont.Style:= [fsBold];
      end;
      {
      if ACol = C_STA then begin
          ABrush.Color:=clYellow;
          if gdIf.Cells[ACol, ARow] <> 'C' then
              AFont.Color:= clRed;
      end;
      }

      if ACol in [0..C_STA] then begin
          sState:= gdIf.Cells[C_STA, ARow];
          if sState = '' then
             ABrush.Color:= clWhite
          else
          if (sState = '??????') or (sState = '????????') then
              ABrush.Color:= clYellow
          else
          if sState = '????????' then
              ABrush.Color:=$00FFFFE8
          else
          if sState = '????????' then
              ABrush.Color:= $00F5ECFF
          else
              ABrush.Color:=$00D0FFD0
      end;

      if ACol in [C_STA+1 .. nCol -1] then begin
          if ACol <= nCol then begin
              sNull:= gdIf.Cells[ACol, ARow];
              if sNull = '' then
                  ABrush.Color:= $00E1E1E1
              else begin
                  ABrush.Color:= clWhite;
                  if sNull = 'Y' then
                      AFont.Color:= $00E1E1E1
                  else
                      AFont.Color:= clBlack;
              end;
          end;
      end;

  end;
end;

procedure TF_Main.ComPort1RxChar(Sender: TObject; Count: Integer);
var
  cStr:string;
begin
  ComPort1.ReadStr(cStr, Count);
  DoH7180(cStr);
end;

procedure TF_Main.AddDebugLine(cStr:string);
begin
  mmLog.Lines.Add(cStr);
end;

procedure TF_Main.ClearCommData;
begin
  mmTemp.Clear;
end;

procedure TF_Main.SendOrderFrame(DataStr:string);
var
  H7180: TH7180If;
  Bcd, sSender:string;
begin

  TGlobal.DataLog:= 'H7180-> '+DataStr; //Log

  Bcd:= Trim(Copy(DataStr, 15, BarCodeLen));
  if Bcd='' then begin
      SendSPMFrame;
      exit;
  end
  else begin
      //Order
      H7180:= TH7180If.Create;
      try
          H7180.FExamDate:= FormatDateTime('yyyymmdd', now);
          //H7180.FExamSeq := PadLeftStr(IntToStr(DM.GetExamSeq(H7180.FExamDate)), '0', 3);
          H7180.BarCode:= Bcd;
          H7180.DownLoadOrder;
          DM.SaveMaster(H7180);
          if H7180.FOrdState = 'Y' then
              DM.SaveOrderList(H7180);

          DispOnePat_Order(H7180);
          DispOrderList(H7180);

          sSender:= Copy(DataStr,1,41)+' 87';
          sSender:= sSender + H7180.MakeOrderStr;

          TGlobal.DataLog:= 'IF_PC-> '+STX + sSender + ETX + H747CheckSum(sSender) + CR; //Log

          SendStr(STX + sSender + ETX + H747CheckSum(sSender) + CR);

      finally
          H7180.Free;
      end;
  end;
end;

function TF_Main.SendSPMFrame:string;
var
  s:string;
begin
  s:= STX + '>' + ETX + H747CheckSum('>') + CR;
  Result:= s;

  SendStr(s);
end;

procedure TF_Main.SendStr(cStr:string);
begin
  AddDebugLine('INTE->'+cStr);

  Delay(200);
  ComPort1.WriteStr(cStr);
end;

procedure TF_Main.SetCommData(const cStr: string);
begin
  AddDebugLine('H7170->'+cStr);

  mmTemp.Text:= cStr;

  //TGlobal.DataLog:= '[ABL]->'+cStr;
end;

procedure TF_Main.N1_1Click(Sender: TObject);
begin
  if F_CodeSet = nil then
      F_CodeSet:= TF_CodeSet.Create(Self);

  F_CodeSet.ShowModal;
end;

procedure TF_Main.N1_4Click(Sender: TObject);
begin
  PortSetup;
end;

procedure TF_Main.btnCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TF_Main.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  if MessageDlg('?????Ͻðڽ??ϱ?? ', mtWarning, mbOKCancel, 0) <> mrOk then
      CanClose:= False
  else
      CanClose:= True;
end;

procedure TF_Main.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  TGlobal.MainTop:= Self.Top;
  TGlobal.MainLeft:= Self.Left;
  TGlobal.MainWidth:= Self.Width;
  TGlobal.MainHeigh:= Self.Height;
end;


procedure TF_Main.N3Click(Sender: TObject);
begin
  Close;
end;

function TF_Main.GetGridIndex(var TGrid:TAdvStringGrid; ExamDate, ExamSeq:string):integer;
var
  i:integer;
  cDateTime:string;
  cSeq:string;
begin
  Result:= 0;
  if (TGrid.RowCount = 2) and (TGrid.Cells[1,1]='') then begin
      Result:=1;  exit;
  end;

  for i:=1 to TGrid.RowCount -1 do begin
      cDateTime:= Copy(Trim(TGrid.Cells[1, i]),1,10);
      cDateTime:= Copy(cDateTime,1,4) + Copy(cDateTime,6,2) + Copy(cDateTime,9,2);
      cSeq:= Trim(TGrid.Cells[2,i]);

      if (cDateTime = ExamDate) and (ExamSeq = cSeq) then
      begin
          Result:= i;
          exit;
      end;
  end;

  TGrid.AddRow;
  Result:= TGrid.RowCount -1;

end;

procedure TF_Main.GridSetting;
var
  i,K:integer;
begin
  //??Ÿ?д? üũ?Ϸ??? ABBR+1 ????.
  K:= C_STA;
  for i:=0 to TCode.AbbrList.Count -1 do begin
      INC(K);
      gdIf.InsertCols(K,1);
      gdIf.Cells[K,0]:= TCode.AbbrList.Strings[i];
  end;

  //gdIf.HideColumns(C_ADT, C_ANO);

end;

procedure TF_Main.gdIfClickCell(Sender: TObject; ARow, ACol: Integer);
var
  BCD, AN, PID, PNM,
  Abbr,Res:string;
  i:integer;
begin
  if ( ARow > 0 ) then begin
      ClearResultGrid;
      edOld.Text:= gdIf.Cells[C_BCD, ARow];
      lbRow.Caption:= IntToStr(ARow);
      //pnBCD.Caption    := gdIf.Cells[C_BCD, ARow];
      //pnAcptNo.Caption := gdIf.Cells[C_ANO, ARow];
      //pnPID.Caption    := gdIf.Cells[C_PID, ARow];
      //pnPNM.Caption    := gdIf.Cells[C_PNM, ARow];
      for i:= C_STA+1 to gdIf.AllColCount -1 do begin
          Abbr:= gdIf.Cells[i, 0];
          Res := gdIf.Cells[i, ARow];
          if Res <> '' then
              AddResultGrid(Abbr, Res);
      end;
  end;
end;

procedure TF_Main.AddResultGrid(Abbr, Res: string);
var
  i, j:integer;
begin
  i:=1;
  while (i < gdResult.ColCount) do begin
      for j:=1 to gdResult.RowCount -1 do begin
          if gdResult.Cells[i, j] = '' then begin
              gdResult.Cells[i, j]:= abbr;
              gdResult.Cells[i+1, j]:= Res;
              exit;
          end;
      end;
      i:= i+2;
  end;
end;

procedure TF_Main.ClearResultGrid;
begin
  gdResult.ClearNormalCells;
end;

procedure TF_Main.AddViewLog(Str: string);
begin
  mmView.Lines.Add(Str);
end;

procedure TF_Main.ChangeUpState(BarCode, UpState:string);
var
  nRow:integer;
begin
  nRow:= FindGridIndex(gdIf, BarCode);

  if nRow = 0 then exit;

  //?̹? ???ۿϷ??ΰ??? ???????? ????.
  if gdIf.Cells[C_STA, nRow] <> '???ۿϷ?' then begin
      if UpState = 'Y' then
          gdIf.Cells[C_STA, nRow] := '???ۿϷ?'
      else
          gdIf.Cells[C_STA, nRow] := '??????';
  end;

  gdIf.Refresh;
  gdIf.OnClickCell(nil, nRow, C_BCD);

end;

procedure TF_Main.gdIfCanEditCell(Sender: TObject; ARow, ACol: Integer;
  var CanEdit: Boolean);
begin
  if ARow > 0 then begin
      if ACol in [0] then
          CanEdit:= True
      else
          CanEdit:= False;
  end;
end;

procedure TF_Main.N6Click(Sender: TObject);
var
  i, UpCount:integer;
  ErrMsg:string;
  bCheck:boolean;
  TMaster:TH7180If;
begin

  if Not CheckBoxCheckYN(gdIf) then
      exit;

  if MessageDlg('?????? ?˻??????? ?????Ͻðڽ??ϱ??', mtConfirmation, mbOKCancel, 1) <> mrOk then
      exit;

  UpCount:= 0;
  for i:=1 to gdIf.RowCount -1 do begin
      bCheck:= False;
      gdIf.GetCheckBoxState(0, i, bCheck);
      if bCheck then begin
          TMaster:= TH7180If.Create;
          try
              TMaster.BarCode:= gdIf.Cells[C_BCD, i];
              TMaster.FExamDate:= FormatDateTime('yyyymmdd', now);
              if DM.DownLoadOrder_Result(TMaster) then begin
                  gdIf.Cells[C_PID, i]:= TMaster.FPatId;
                  gdIf.Cells[C_PNM, i]:= TMaster.FPatNm;

                  if DM.UpLoadResult(TMaster.BarCode, ErrMsg) then begin
                      Inc(Upcount);
                      DM.ChangeState(TMaster.BarCode, 'Y');
                      gdIf.Cells[C_STA, i]:= '???ۿϷ?';
                      gdIf.SetCheckBoxState(0, i, False);
                      gdIf.Refresh;
                      gdIf.OnClickCell(nil, i, C_BCD);
                  end
                  else begin
                      if ErrMsg <> '' then
                          AddViewLog('???ڵ?:'+TMaster.BarCode+' ???? ???? -> MSG['+ErrMsg+']');
                  end;
              end
              else begin
                  AddViewLog('???ڵ?:'+TMaster.BarCode+' -> ?????? ?????ϴ?! ???ڵ带 Ȯ???ϼ???');
                  Continue;
              end;
          finally
              TMaster.Free;
          end;
      end;
  end;

  ShowMessage(IntToStr(UpCount)+' ?? ???ۿϷ?!');
end;

procedure TF_Main.mmViewDblClick(Sender: TObject);
begin
  mmView.Clear;
end;

procedure TF_Main.gdResultGetCellColor(Sender: TObject; ARow,
  ACol: Integer; AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
var
  S:string;
begin
  if ARow > 0 then begin
      if ACol in [1,3,5] then begin
          if gdResult.Cells[ACol, ARow] <> '' then
          ABrush.Color:= $00FFCAE4;
      end;

      if ACol in [2,4,6] then begin
          S:= gdResult.Cells[ACol, ARow];
          if S = '' then begin
              ABrush.Color:= $00E1E1E1;
          end
          else begin
              if S = 'Y' then begin
                  ABrush.Color:= clWhite;
                  AFont.Color:= $00E1E1E1;
              end
              else begin
                  AFont.Color:= clBlack;
                  ABrush.Color:= $00FFCAE4;
              end;
          end;
      end;
  end;
end;

procedure TF_Main.DispOnePatInfo(TMaster: TH7180If);
var
  nRow:integer;
begin
  nRow:= GetGridIndex(gdIf, TMaster.BarCode);

  if nRow = 0 then exit;
  
  gdIf.AddCheckBox(0, nRow, False, False);

  with TMaster do begin
      gdIf.Cells[C_BCD, nRow] := BarCode;
      //gdIf.Cells[C_ANO, nRow] := FAcptNo;
      gdIf.Cells[C_PID, nRow] := FPatId;
      gdIf.Cells[C_PNM, nRow] := FPatNm;
  end;

  //gdIf.OnClickCell(nil, nRow, C_DTM);
end;

procedure TF_Main.CLEAR1Click(Sender: TObject);
begin
  if MessageDlg('ȭ?????? ?????͸? ???? ?????Ͻðڽ??ϱ??', mtConfirmation, mbOKCancel, 1) <> mrOK then exit;

  gdIf.ClearNormalCells;
  gdIf.RowCount:=2;

  gdIf.Refresh;

end;

procedure TF_Main.DoH7180(const cData: string);
var
 i:integer;
 cBuffer,cStr:string;
begin
     // ?????????? ?????? ?ڷḦ ????
     cBuffer:=vcTxData+cData;
     for i:= 1 to Length(cBuffer) do begin
       case cBuffer[i] of
           SOH: ;//AddDebugLine('H7180:SOH');
           STX:begin
                   cStr:='';
               end;
           ETB: ;//AddDebugLine('H7180:[STX]'+cStr+'[ETB]');
           ETX:begin
                   //AddDebugLine('H7180:[STX]'+cStr+'[ETX]');
                   SetCommData(cStr);
                   H7180_Rcv;
                   vcTxData:='';
                   cStr:='';
                   ClearCommData;
               end;
           EOT: ;//AddDebugLine('H7180:EOT');
           else
              cStr:=cStr+cBuffer[i];
       end;
     end;
     vcTxData:=cStr;
end;

function TF_Main.H747Str(cData:string): string;
begin
   H747Str := STX + cData + ETX + H747CheckSum(cData) + CR;
end;

function TF_Main.LoadComPort: boolean;
begin
  Result:= False;
  pnPort.Visible:= False;

  { CPort }
  ComPort1.LoadSettings(stIniFile, TGlobal.AppPath+IniFileName);
  ComPort1.Open;
  Result:= True

  { MSComm }
  {try
      TGlobal.ComPortIniLoad;
      MSComm1.CommPort := TGlobal.ComPortSet.PortNum;
      MSComm1.Settings := TGlobal.ComPortSet.Settings;
      MsComm1.DTREnable:= TGlobal.ComPortSet.Dtr;
      MsComm1.RTSEnable:= TGlobal.ComPortSet.Rts;
      MsComm1.Handshaking:= TGlobal.ComPortSet.HandShake;
  except
      on e:Exception do begin
          ShowMessage(e.Message);
          exit;
      end;

  end;
  Result:= True;
  }

end;

function TF_Main.H747CheckSum(SData:string): string;
var
	Sum, i : LongInt;
begin
   Sum := 0;
   for i := 1 to Length(SData) do
		Sum := Sum + Ord(SData[i]);
   H747CheckSum := Copy(Format('%4x', [Sum]), 3, 2);
end;


procedure TF_Main.H7180_Rcv;
var
  Temp:string;
  Frame, Func:string;
  BarCode, ResStr, IfCode, sResult,Flag:string;
  nCnt, j:integer;
  H7180: TH7180If;
  sSender, sSend:string;
  TResultThread: TMyThread;
begin
  Temp:= mmTemp.Text;

  Frame := copy(Temp, 1, 1);
  Func  := UpperCase(copy(Temp, 2, 2));

  if Frame = '' then Exit;

  case Frame[1]  of
      ':',
      '1',
      '2': begin
               if (UpperCase(Func)='AB') or
                  (UpperCase(Func)='D1') or
                  (UpperCase(Func)='N1') or
                  (UpperCase(Func)='A1') then
               begin
                   if USE_THREAD then begin
                       ResultString:= Temp;
                       TResultThread:= TMyThread.Create(False);
                       Delay(500);  //?ڲ? ?ι? ??????.. ?ʹ????? ???? ?׷???..
                       sSend:= SendSPMFrame;
                       TGlobal.DataLog:= 'IF_PC-> '+sSend; //Log
                   end
                   else begin
                       H7180_ResultProcess(Temp);
                       sSend:= SendSPMFrame;
                       TGlobal.DataLog:= 'IF_PC-> '+sSend; //Log
                   end;

               end
               else begin
                   TGlobal.DataLog:= 'H7180-> '+Temp; //Log
                   sSender:= SendSPMFrame;
                   TGlobal.DataLog:= 'IN_PC-> '+sSender;
               end;
           end;
      ';': SendOrderFrame(Temp);
      else
          SendSPMFrame;
  end;
end;

procedure TF_Main.H7180_ResultProcess(sData: string);
var
  H7180: TH7180If;
  ResStr:string;
  ResCnt, i:integer;
  sIfCode,
  Flag,
  Bcd, Abbr:string;
  SvrMsg, sSend:string;
begin
  if sData = '' then exit;

  TGlobal.DataLog:= 'H7180-> '+sData; //Log

  H7180:= TH7180If.Create;
  try
      Bcd:= Trim(Copy(sData,15, BarCodeLen));

      H7180.BarCode  := Bcd;
      H7180.FExamDate:= FormatDateTime('yyyymmdd', now);

      ResCnt := StrToIntDef(Trim(Copy(sData,49,2)),0);
      if ResCnt=0 then begin
          Exit;
      end;

      //????üũ
      if DM.SelectLocalOrder(H7180.BarCode) = '' then begin
          H7180.DownLoadOrder;
          DM.SaveMaster(H7180);
          if H7180.FOrdState = 'Y' then
              DM.SaveOrderList(H7180);
      end;

      DispOnePat_Result(h7180);

      ResStr:=Copy(sData,52,300);

      for i:=0 to ResCnt -1 do begin
          sIfCode:= Trim((Copy(ResStr,i*10+1,3)));

          h7180.slResIfCode.Add(sIfCode);
          h7180.slResult.Add(Trim(Copy(ResStr,i*10+3,6)));
          h7180.slResExCode.Add(DM.GetExamCode(sIfCode));
      end;

      if ResCnt > 0 then begin
          DispResultList(h7180);
          DM.SaveResultList(h7180);

          if DM.UpLoadResult(h7180.BarCode, SvrMsg) then begin
              DM.ChangeState(h7180.BarCode, 'Y');
              ChangeUpState(h7180.BarCode, 'Y');
          end
          else begin
              ChangeUpState(h7180.BarCode, '');
              if SvrMsg <> '' then begin
                  AddViewLog('???ڵ?:'+H7180.BarCode+' ???? ???? -> MSG['+SvrMsg+']');
                  TGlobal.LogMsg:= '???ڵ?:'+H7180.BarCode+' ???? ???? -> MSG['+SvrMsg+']';  //Err.log
              end;
          end;

      end;

  finally
      H7180.Free;
  end;

end;


{ EPortError }

constructor EPortError.Create;
begin
  ShowMessage('Port Error!');
end;

function TF_Main.GetGridIndex(TGrid: TAdvStringGrid;
  BarCode: string): integer;
var
  i:integer;
  cDateTime:string;
  cSeq:string;
begin
  Result:= 0;
  if (TGrid.RowCount = 2) and (TGrid.Cells[1,1]='') then begin
      Result:=1;  exit;
  end;

  for i:=1 to TGrid.RowCount -1 do begin
      if TGrid.Cells[C_BCD, i] = BarCode then begin
          Result:= i;
          exit;
      end;
  end;

  TGrid.AddRow;
  Result:= TGrid.RowCount -1;

end;

procedure TF_Main.DispOrderList(TMaster: TH7180If);
var
  i, nRow, nCol: integer;
  Abbr:string;
begin
  nRow:= GetGridIndex(gdIf, TMaster.BarCode);
  if nRow = 0 then exit;

  for i:=0 to TMaster.slExCode.Count -1 do begin
      Abbr:= DM.GetAbbr(TMaster.slExCode.Strings[i]);
      if Abbr <> '' then begin
          nCol:= GetAbbrIndex(Abbr);
          if nCol = 0 then
              Continue
          else
              gdIf.Cells[nCol, nRow]:= 'Y';
      end;
  end;
end;

procedure TF_Main.DispOnePat_Result(TMaster: TH7180If);
var
  nRow:integer;
begin
  nRow:= GetGridIndex(gdIf, TMaster.BarCode);

  if nRow > 0 then begin
      gdIf.AddCheckBox(0, nRow, False, False);
      with TMaster do begin
          //gdIf.Cells[C_SEQ, nRow] := FExamSeq;
          //gdIf.Cells[C_DTM, nRow] := ViewDateTime(FExamDate+FExamTime);
          gdIf.Cells[C_BCD, nRow] := BarCode;
          if FPatId <> '' then
              gdIf.Cells[C_PID, nRow] := FPatId;
          if FPatNm <> '' then
              gdIf.Cells[C_PNM, nRow] := FPatNm;

          //gdIf.Cells[C_ADT, nRow] := FAcptDt;
          //gdIf.Cells[C_ANO, nRow] := FAcptNo;
          if gdIf.Cells[C_STA, nRow] <> '???ۿϷ?' then
              gdIf.Cells[C_STA, nRow] := '????????'
      end;
  end;

end;

procedure TF_Main.DispResultList(TMaster: TH7180If);
var
  i, nRow, nCol: integer;
  Abbr:string;
begin
  nRow:= FindGridIndex(gdIf, TMaster.BarCode);
  if nRow = 0 then exit;

  for i:=0 to TMaster.slResIfCode.Count -1 do begin
      Abbr:= DM.GetAbbr(TMaster.slResExCode.Strings[i]);
      if Abbr <> '' then begin
          nCol:= GetAbbrIndex(Abbr);
          if nCol = 0 then
              Continue
          else
              if TMaster.slResult.Strings[i] <> '' then
                  gdIf.Cells[nCol, nRow]:= TMaster.slResult.Strings[i]
              else begin
                  if gdIf.Cells[nCol, nRow] <> 'Y' then
                      gdIf.Cells[nCol, nRow]:= ' ';
              end;
      end;
  end;

end;

function TF_Main.FindGridIndex(TGrid: TAdvStringGrid;
  BarCode: string): integer;
var
  i:integer;
  cDateTime:string;
  cSeq:string;
begin
  Result:= 0;
  if (TGrid.RowCount = 2) and (TGrid.Cells[1,1]='') then begin
      Result:=1;  exit;
  end;

  for i:=1 to TGrid.RowCount -1 do begin
      if TGrid.Cells[C_BCD, i] = BarCode then begin
          Result:= i;
          exit;
      end;
  end;
end;

procedure TF_Main.N7Click(Sender: TObject);
begin
  if F_Test = nil then
      F_Test:= TF_Test.Create(Self);

  F_Test.Show;
end;

procedure TF_Main.Rcv1Click(Sender: TObject);
var
  S:string;
begin
  S:= InputBox('TEST', '??????????', '');
  if S <> '' then
      DoH7180(S);
end;

{ TMyThread }

constructor TMyThread.Create(CreateSuspended: Boolean);
begin
  inherited Create(CreateSuspended);
  FreeOnTerminate:= True;
end;

procedure TMyThread.Execute;
begin
  inherited;
  Synchronize(Run_ResultProcess);
end;

procedure TMyThread.Run_ResultProcess;
var
  S:String;
begin
  S:= Copy(ResultString,1,Length(ResultString));
  F_Main.H7180_ResultProcess(S);
end;

procedure TF_Main.gdIfDblClickCell(Sender: TObject; ARow, ACol: Integer);
begin
  if ARow > 0 then begin
      if pnBcd.Visible = False then
          pnBcd.Visible:= True;
  end;
end;

procedure TF_Main.btnBcdCloseClick(Sender: TObject);
begin
  pnBcd.Visible:= False;
end;

procedure TF_Main.btnBcdChangeClick(Sender: TObject);
var
  OldBcd, NewBcd:string;
begin
  OldBcd:= Trim(edOld.Text);
  NewBcd:= Trim(edNew.Text);
  if OldBcd = NewBcd then begin
      ShowMessage('?????? ??ȣ?? ?????Ҽ? ?????ϴ?!');
      exit;
  end;

  if NewBcd = '' then begin
      ShowMessage('?????Ͻ? ??ȣ?? ?Է??ϼ???!');
      edNew.SetFocus;
      exit;
  end;

  DM.ChangeBarCode(OldBcd, NewBcd);

  gdIf.Cells[C_BCD, StrToInt(lbRow.Caption)]:= edNew.Text;
  
end;

procedure TF_Main.Panel6MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
  ReleaseCapture;
  SendMessage(pnBcd.Handle, WM_SYSCOMMAND, 61458, 0);
end;

procedure TF_Main.pnBcdMouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
  ReleaseCapture;
  SendMessage(pnBcd.Handle, WM_SYSCOMMAND, 61458, 0);
end;

initialization
  vcTxData:='';

  TGlobal:= TGlobalVar.Create;
  TConnection:= TDbConnection.Create;
  TCode:= TCodeInfo.Create;


finalization


  TCode.Free;
  TConnection.Free;
  TGlobal.Free;

end.
