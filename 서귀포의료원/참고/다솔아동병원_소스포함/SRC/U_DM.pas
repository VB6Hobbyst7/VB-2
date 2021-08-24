unit U_DM;

interface

uses
  SysUtils, Windows, Classes, DB, ADODB, Forms, U_Main, Dialogs, U_IFClass,
  DBTables, math;

type
  TDM = class(TDataModule)
    qryC1: TADOQuery;
    qryC2: TADOQuery;
    qryC3: TADOQuery;
    qrySUp: TADOQuery;
    qryUp1: TADOQuery;
    qryUpOne: TADOQuery;
    qrySUp1: TADOQuery;
    qryV: TADOQuery;
    spUp: TADOStoredProc;
    conHosp: TADOConnection;
    qrySOrder: TADOQuery;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    function UploadHost_State_DSW(ADT, ANO, PID, STA: string): boolean;
  public
    procedure DeleteMaster_One(ExamDate, ExamSeq: string);
    procedure DeleteResult_OneSeq(ExamDate, ExamSeq: string);
    procedure DeleteOldData(nDays: Cardinal =60);
    function CheckSetCode(ExamCode:string):boolean;
    function CheckSetCode_ORD(OrdCode:string):boolean;

    procedure SaveMaster(TMaster:TIfMaster);

    {사용하지 않음}
    function SaveOrder(TMaster:TIfMaster):string;

    function SaveResult(TMaster:TIfMaster): boolean;
    function GetExamSeq(ExamDate:string):string;
    function GetExamSeq_QC(ExamDate:string):string;

    procedure UpdateOrderState(ExamDate:string; ExamSeq, nState:integer);
    procedure DeleteData(ExamDate, ExamSeq:string);

    function GetExamCode(OrdCode, IfCode:string):string;
    function GetExamPanelCode(PCode, IfCode:string):string;
    function CheckOrderCode(OrdCode:string):boolean;
    function getOrdCode(PanelName:string):string;
    function GetExamUpCode(IfCode:string):Variant;
    function UpdateSpcid(ExamDate,cSpcid:string; ExamSeq:integer):boolean;
    function GetQcBarCode(cLot, cIName:string):string;

    function SaveOneCode(OCode, sPanel, Ecode,EName,Abbr,IfCd, UpCode,RefL,RefH,Seq:string):boolean;
    function DeleteOneCode(OrdCode, ExamCode:string):boolean;

    //Inst
    function SaveOneInst(ICode,IName,HCode,Loc,Seq, POCT:string):boolean;
    function DeleteOneInst(ICode:string):boolean;

    //Panel
    function SaveOnePanelCode(PCode,PName,Seq:string):boolean;
    function DeleteOnePanelCode(PCode:string):boolean;
    function SetCodePanel(PCode,ECode,Seq:string):boolean;
    function DelCodePanel(PCode,ECode:string):boolean;

    //Flag Set
    function SaveOneFlagSet(ICode, Flag, PCode, Seq: string):boolean;
    function DeleteOneFlagSet(ICode, PCode, Flag:string):boolean;

    procedure SendOutMessage(sIp, sState, sSampleId:string);
    procedure SendDeltaMessage(sIp, sState, sSampleId:string);

    procedure SaveState(ExamDate, ExamSeq, ErrMsg, UpState: string);
    function SaveSpcId(Spcid, ExamDate, ExamSeq: string):boolean;
    function DeleteOneData(ExamDate, ExamSeq:string):boolean;

    function UploadResult(TMaster:TIfMaster):boolean;
    function UploadHost_One_YJ(ReceiptNo, OCD, ECD, ResVal:string):boolean;
    function UploadState_YJ(ReceiptNo, OrdCode, UpState:string):boolean;
    function DownLoadOrder_YJ(TMaster:TIfMaster):boolean;
    function DownLoadOrder_PTM(TMaster:TIfMaster):boolean;

    function UploadHosp_One_DJ(TMaster:TIfMaster):boolean;
    function UploadState_DJ(TMaster:TIfMaster):boolean;
    function DownLoadOrder_DJ(TMaster:TIfMaster):boolean;

    function DownLoadOrder_DJI_WORK(FDT:string):string;
    function DownLoadOrder_DJI(TMaster:TIfMaster):boolean;
    function UploadHosp_One_DJI(TMaster:TIfMaster):boolean;

    function DownLoadOrder_SCHUH(TMaster:TIfMaster):boolean;
    function DownLoadOrder_SCHUH_One(TMaster:TIfMaster):boolean;
    function UploadHosp_One_SCHUH(TMaster:TIfMaster):boolean;

    function DownLoadOrder_CBD(TMaster:TIfMaster):boolean;
    function UploadHosp_One_CBD(TMaster:TIfMaster):boolean;
    function UploadHosp_STATE_CBD(BCD, OCD:string):boolean;

    function DownLoadOrder_JSD(TMaster:TIfMaster):boolean;
    function UploadHost_One_JSD(BCD, ECD, IfCd, RES:string):boolean;

    //제일병원
    function MakeBarCode(BCD:string):string;
    function GetInstNo(INM, ICD:string):integer;
    function UploadResult_Direct(TMaster:TIfMaster):boolean;

    function DownLoadOrder_JEIL(TMaster:TIfMaster):boolean;
    //결과 Table
    function UploadHospOne_JEIL_RES(TMaster:TIfMaster):boolean;
    function UploadHospOne_JEIL_RMK(TMaster:TIfMaster):boolean;
    //검사실 처방 Table           //341, 131
    //처방 Table(외래)
    //처방 Table(입원)
    //검사처방별 검사장비코드 Table

    function DownLoadOrder_QC_JEIL(TMaster:TIfMaster):boolean;
    function UploadHospOne_QC_JEIL(ADT, BCD, ECD, RES, RMK: string; OSEQ:integer):boolean;
    //
    function DownLoadOrder_JAIN(TMaster:TIfMaster):boolean;
    function DownLoadPAT_JAIN(TMaster:TIfMaster):boolean;
    function UploadHost_One_JAIN(BCD, ECD, ResVal:string):boolean;
    function UploadHost_State_JAIN(BCD, STA:string):boolean;

    function DownLoadOrder_DSW(TMaster:TIfMaster):boolean;
    function DownLoadOrder_KY(TMaster:TIfMaster):boolean;
    function DownLoadOrder_KY_BCD(TMaster:TIfMaster):boolean;
    function UploadHost_One_PTM(BCD, ECD, ResVal:string):boolean;
    function UploadHost_One_DSW(BCD, ECD, ResVal, LH:string; var ErrMsg:string):boolean;
    function UploadHost_One_KY(ADT, SLP, ANO, ECD, PID, RES:string; var ErrMsg:string):boolean;
    function UploadHost_State_KY(ADT, SLP, ANO, PID, STA:string):boolean;
    function GetANO(BCD, ECD:string):string;
    function GetExamCode_Var(IfCode:string):Variant;

    function DownLoadOrder_JND_RES(TMaster:TIfMaster):boolean;
    function UploadHosp_One_JND(TMaster:TIfMaster):boolean;


    function CheckLot(INm, Enm, Lot, Typ, Lev:string):boolean;
    procedure SaveOneLotInfo(INm, LotNm, Lev, ENM, Mean, SD, Fdt, Tdt, sLow, sHigh, Typ:string);

    procedure SetExamLotMean(TMaster:TIfMaster);
    procedure SaveQC(TMaster:TIfMaster);

    procedure SetDownCode(TMaster:TIfMaster);
  end;

var
  DM: TDM;

implementation

uses SetDataBase, GlobalVar, Variants, StringLib, U_CodeInfo;

{$R *.dfm}

procedure TDM.DataModuleCreate(Sender: TObject);
begin
  DeleteOldData(90); //3개월..
  TConnection.LocalCon.Close;
  TGlobal.LocalMDBCompress('SANSOFT');

  {LOCAL}
  qryUp1.Connection:= TConnection.LocalCon;
  qryUpOne.Connection:= TConnection.LocalCon;
  qryC1.Connection:= TConnection.LocalCon;
  qryC2.Connection:= TConnection.LocalCon;
  qryC3.Connection:= TConnection.LocalCon;
  qryV.Connection:= TConnection.LocalCon;

  //TuxInit;

  {HOSP}
  spUp.Connection:= TConnection.hospCon;
  qrySOrder.Connection  := TConnection.hospCon;
  qrySUp.Connection    := TConnection.HospCon;
  qrySUp1.Connection  := TConnection.hospCon;
  spUp.Connection     := TConnection.hospCon;
end;

procedure TDM.DeleteData(ExamDate, ExamSeq:string);
begin
  DeleteResult_OneSeq(ExamDate,ExamSeq);
  DeleteMaster_One(ExamDate, ExamSeq);
end;

procedure TDM.DeleteOldData(nDays: Cardinal=60);
var
  tSql: TQueryInfo;
  cDays: string;
begin
  cDays:= FormatDateTime('yyyymmdd', now - nDays);
  tSql:= TQueryInfo.Create;

  try
    with tSql do begin
        Clear;
        AddSql(' Delete From TB_Result Where ExamDate <= '''+cDays+''' ');
        LocalExcute;

        Clear;
        AddSql(' Delete From TB_Master Where ExamDate <= '''+cDays+''' ');
        LocalExcute;
    end;

  finally
    tSql.Free;
  end;

end;

function TDM.DeleteOneCode(OrdCode, ExamCode: string): boolean;
var
  TSql: TQueryInfo;
begin
  TSql:= TqueryInfo.Create;
  try
      with TSql do begin
          Clear;
          AddSql(' Delete From TB_CodeInfo ');
          AddSql(' Where OrdCode = '''+OrdCode+''' ');
          AddSql('   And ExamCode = '''+ExamCode+''' ');
          Result:= LocalExcute;
      end;

  finally
      TSql.Free;
  end;

end;

procedure TDM.DeleteMaster_One(ExamDate, ExamSeq: string);
var
  TSql: TQueryInfo;
begin
  TSql:= TQueryInfo.Create;
  try
      with TSql do
      begin
          AddSql(' Delete From TB_Master  ');
          AddSql(' Where ExamDate = '''+ExamDate+''' ');
          AddSql('   And ExamSeq  = '''+ExamSeq+''' ');
          LocalExcute;
      end;

  finally
      TSql.Free;
  end;

end;

procedure TDM.DeleteResult_OneSeq(ExamDate, ExamSeq: string);
var
  TSql: TQueryInfo;
begin
  TSql:= TQueryInfo.Create;
  try
      with TSql do
      begin
          AddSql(' Delete From TB_Result  ');
          AddSql(' Where ExamDate = '''+ExamDate+''' ');
          AddSql('   And ExamSeq  = '''+ExamSeq+''' ');
          LocalExcute;
      end;

  finally
      TSql.Free;
  end;

end;

function TDM.GetExamSeq(ExamDate: string):string;
var
  TSql: TQueryInfo;
  QryEx: TAdoQuery;
  Seq:integer;
begin
  Result:= '001';
  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Application);

  try
      with TSql do begin
        Clear;
        AddSql(' Select iif( isNull(Max(Val(ExamSeq))), 1, Max(Val(ExamSeq))+1) As SEQ From TB_Master ');
        AddSql(' Where ExamDate = '''+ExamDate+'''         ');
        RCount:= LocalSelect(QryEx);

        if RCount > 0 then begin
            Seq:= QryEx.Fields[0].AsInteger;
            Result:= PadLeftStr(IntToStr(Seq), '0', 3);
        end;
    end;

  finally
      TSql.Free;
      QryEx.Free;
  end;
end;

function TDM.GetQcBarCode(cLot, cIName: string): string;
var
  TSql:TQueryInfo;
  QryEx:TAdoQuery;
begin
  Result:='';
  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Application);

  try
      with TSql do begin
          Clear;
          AddSql(' Select BarCode From QCID ');
          AddSql(' Where Loc = '''+cIName+''' ');
          AddSql('   And Lot = '''+cLot+''' ');
          if LocalSelect(QryEx) > 0 then
              Result:= QryEx.Fields[0].AsString;
      end;
  finally
      QryEx.Free;
      TSql.Free;
  end;

end;

function TDM.SaveOneCode(OCode, sPanel, Ecode,EName,Abbr,IfCd, UpCode,RefL,RefH,Seq:string): boolean;
var
  i:integer;
begin
  Result:= False;

  with DM.qryC1 do begin
      Close;
      SQL.Text:= ' Select * From TB_CodeInfo '+
                 ' Where OrdCode = '''+OCode+''' '+
                 '   And ExamCode = '''+ECode+''' ';
      Open;

      if RecordCount > 0 then begin
          Close;

          SQL.Text:=' Update TB_CodeInfo Set '+
                    '     Panel = '''+sPanel+''' '+
                    '   , IfCode =  :IfCd '+
                    '   , UpCode =  :UpCode '+
                    '   , ExamName = '''+EName+''' '+
                    '   , Abbr =  '''+Abbr+''' '+
                    '   , RefLow =  :RLow '+
                    '   , RefHigh = :RHigh '+
                    '   , DispSeq = :DispSeq '+
                    ' Where OrdCode = '''+OCode+''' '+
                    '   And ExamCode = '''+ECode+''' ';
      end
      else begin
          Close;
          SQL.Text:=' Insert Into TB_CodeInfo '+
                    ' (OrdCode, Panel, ExamCode, IfCode, UpCode,ExamName, Abbr,RefLow,RefHigh,DispSeq) '+
                    ' Values '+
                    ' ('''+OCode+''', '''+sPanel+''', '''+ECode+''',:IfCd, :UpCode, '''+EName+''', '''+Abbr+''' '+
                    ' ,:RLow, :RHigh, :DispSeq)  ';
      end;

      Parameters.ParamByName('IfCd').Value:= IfCd;
      Parameters.ParamByName('UpCode').Value:= UpCode;
      Parameters.ParamByName('RLow').Value:= StrToFloatDef(RefL,0);
      Parameters.ParamByName('RHigh').Value:= StrToFloatDef(RefH,0);
      Parameters.ParamByName('DispSeq').Value:= StrToIntDef(Seq,0);

      ExecSql;

      Result:= True;
  end;

end;

function TDM.SaveResult(TMaster:TIfMaster): boolean;
var
  TSql: TQueryInfo;
  QryEx:TADOQuery;
begin
  Result:= False;

  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Self);

  try
      with Tsql do begin
          Clear;
          SqlCmd:= ' Select * FRom TB_Result            '+
                   ' Where ExamDate = '''+TMaster.FExamDate+''' '+
                   '   And ExamSeq  = '''+TMaster.FExamSeq+'''  '+
                   //'   And OrdCODE    = '''+FOrdCode+'''  '+
                   '   And ExamCode = '''+TMaster.FExamCode+''' ';
          RCount:= LocalSelect(QryEx);

          if RCount = 0 then begin
              Clear;
              AddSql( ' Insert Into TB_Result    ');
              AddSql( ' ( ExamDate, ExamSeq, OrdCODE, ExamCode, PNAME, UpCode, Flag, Result ) ');
              AddSql( ' Values                ');
              AddSql( ' (  '''+TMaster.FExamDate+'''   ');
              AddSql( '  , '''+TMaster.FExamSeq+'''    ');
              AddSql( '  , '''+TMaster.FOrdCode+'''   ');
              AddSql( '  , '''+TMaster.FExamCode+'''  ');
              AddSql( '  , '''+TMaster.FExamPanel+''' ');
              AddSql( '  , '''+TMaster.FUpCode+'''    ');
              AddSql( '  , '''+TMaster.FFlag+'''      ');
              AddSql( '  , '''+TMaster.FResult+''')   ');
          end
          else begin
              Clear;
              AddSql( ' Update TB_Result Set   ');
              AddSql( '   Result = '''+TMaster.FResult+''' ');
              AddSql( ' , Flag = '''+TMaster.FFlag+''' ');
              AddSql(' Where ExamDate = '''+TMaster.FExamDate+''' ');
              AddSql('   And ExamSeq  = '''+TMaster.FExamSeq+'''  ');
              AddSql('   And ExamCode = '''+TMaster.FExamCode+''' ');
          end;

          Result:= LocalExcute;
      end;
  finally
      QryEx.Free;
      Tsql.Free;
  end;
end;

function TDM.UpdateSpcid(ExamDate, cSpcid: string; ExamSeq: integer):boolean;
var
  TSql: TQueryInfo;
begin
  Result:= False;

  TSql:= TQueryInfo.Create;
  try
      with TSql do
      begin
          Clear;
          AddSql(' Update TB_Master Set BarCode = '''+cSpcid+''' ');
          AddSql(' Where ExamDate = '''+ExamDate+''' ');
          AddSql('   And ExamSeq = '+IntToStr(ExamSeq)+' ');
          Result:= LocalExcute;
      end;
  finally
      TSql.Free;
  end;
end;

procedure TDM.UpdateOrderState(ExamDate:string; ExamSeq, nState:integer);
var
  TSql: TQueryInfo;
begin
  TSql:= TQueryInfo.Create;
  try
      with TSql do
      begin
          AddSql(' Update TB_Master Set State = '+IntToStr(nState)+' ');
          AddSql(' Where ExamDate = '''+ExamDate+''' ');
          AddSql('   And ExamSeq = '+IntToStr(ExamSeq)+' ');
          LocalExcute;
      end;
  finally
      TSql.Free;
  end;

end;

function TDM.DeleteOneInst(ICode: string): boolean;
var
  TSql: TQueryInfo;
begin
  TSql:= TqueryInfo.Create;
  try
      with TSql do begin
          Clear;
          AddSql(' Delete From TB_Inst ');
          AddSql(' Where ICode = '''+ICode+''' ');

          Result:= LocalExcute;
      end;

  finally
      TSql.Free;
  end;
end;

function TDM.SaveOneInst(ICode, IName, HCode, Loc, Seq, POCT: string): boolean;
var
  TSql: TQueryInfo;
begin
  Result:= False;

  TSql:= TqueryInfo.Create;

  try
      with TSql do begin
          Clear;
          AddSql(' Delete From TB_Inst ');
          AddSql(' Where ICode = '''+ICode+''' ');
          LocalExcute;

          Clear;
          AddSql(' Insert Into TB_Inst ');
          AddSql(' (ICode,IName, HCode,Location,DispSeq,POCT) ');
          AddSql(' Values ');
          AddSql(' ('''+ICode+''', '''+IName+''', '''+HCode+''' ');
          AddSql(' ,'''+Loc+''','+seq+','+POCT+' ) ');
          Result:= LocalExcute;
      end;

  finally
      TSql.Free;
  end;

end;

function TDM.DeleteOnePanelCode(PCode: string): boolean;
var
  TSql: TQueryInfo;
begin
  TSql:= TqueryInfo.Create;
  try
      with TSql do begin
          Clear;
          AddSql(' Delete From TB_Panel_M ');
          AddSql(' Where PCode = '''+PCode+''' ');

          Result:= LocalExcute;
      end;

  finally
      TSql.Free;
  end;

end;

function TDM.SaveOnePanelCode(PCode, PName, Seq: string): boolean;
var
  TSql: TQueryInfo;
begin
  Result:= False;

  TSql:= TqueryInfo.Create;

  try
      with TSql do begin
          Clear;
          AddSql(' Delete From TB_Panel_M ');
          AddSql(' Where PCode = '''+PCode+''' ');
          LocalExcute;

          Clear;
          AddSql(' Insert Into TB_Panel_M ');
          AddSql(' (PCode, PName, DispSeq) ');
          AddSql(' Values ');
          AddSql(' ('''+PCode+''', '''+PName+''', '+seq+' ) ');
          Result:= LocalExcute;
      end;

  finally
      TSql.Free;
  end;

end;

function TDM.DeleteOneFlagSet(ICode, PCode, Flag: string): boolean;
var
  TSql: TQueryInfo;
begin
  TSql:= TqueryInfo.Create;
  try
      with TSql do begin
          Clear;
          AddSql(' Delete From TB_Code_Flag ');
          AddSql(' Where ICode = '''+ICode+''' ');
          AddSql('   And Flag  = '''+Flag+''' ');
          AddSql('   And PCode = '''+PCode+''' ');

          Result:= LocalExcute;
      end;

  finally
      TSql.Free;
  end;

end;

function TDM.SaveOneFlagSet(ICode, Flag, PCode, Seq: string): boolean;
var
  TSql: TQueryInfo;
begin
  Result:= False;

  TSql:= TqueryInfo.Create;

  try
      with TSql do begin
          Clear;
          AddSql(' Delete From TB_Code_Flag ');
          AddSql(' Where ICode = '''+ICode+''' ');
          AddSql('   And Flag  = '''+Flag+''' ');
          AddSql('   And PCode = '''+PCode+''' '); 
          LocalExcute;

          Clear;
          AddSql(' Insert Into TB_Code_Flag ');
          AddSql(' (ICode, Flag, PCode, DispSeq) ');
          AddSql(' Values ');
          AddSql(' ('''+ICode+''', '''+Flag+''', '''+PCode+''', '+seq+' ) ');
          Result:= LocalExcute;
      end;

  finally
      TSql.Free;
  end;

end;

function TDM.DelCodePanel(PCode, ECode: string): boolean;
var
  TSql: TQueryInfo;
begin
  TSql:= TqueryInfo.Create;
  try
      with TSql do begin
          Clear;
          AddSql(' Delete From TB_Panel_D ');
          AddSql(' Where PCode = '''+PCode+''' ');
          AddSql('   And ECode = '''+ECode+''' ');

          Result:= LocalExcute;
      end;

  finally
      TSql.Free;
  end;

end;

function TDM.SetCodePanel(PCode, ECode, Seq: string): boolean;
var
  TSql: TQueryInfo;
begin
  Result:= False;

  TSql:= TqueryInfo.Create;

  try
      with TSql do begin
          Clear;
          AddSql(' Delete From TB_Panel_D ');
          AddSql(' Where PCode = '''+PCode+''' ');
          AddSql('   And ECode = '''+ECode+''' ');
          LocalExcute;

          Clear;
          AddSql(' Insert Into TB_Panel_D ');
          AddSql(' (ECode, PCode, DispSeq) ');
          AddSql(' Values ');
          AddSql(' ('''+ECode+''', '''+PCode+''', '+seq+' ) ');
          Result:= LocalExcute;
      end;

  finally
      TSql.Free;
  end;

end;

procedure TDM.SendOutMessage(sIp, sState, sSampleId: string);
var
  strcmd:string;
begin
  strcmd:= 'net send ' + sIp + ' Error code.  ' + sState + ' SampleNo: ' +sSampleId;
  WinExec(pansichar(strcmd), SW_HIDE);
end;

procedure TDM.SendDeltaMessage(sIp, sState, sSampleId: string);
var
  strcmd:string;
begin
  strcmd:= 'net send ' + sIp + ' Error code.  ' + sState + ' SampleNo: ' +sSampleId + '[ Panic Data!!!! ]';
  WinExec(pansichar(strcmd), SW_HIDE);
end;

procedure TDM.SaveState(ExamDate, ExamSeq, ErrMsg, UpState: string);
var
  TSql:TQueryInfo;
begin
  ErrMsg:= '"'+ErrMsg+'"';

  TSql:= TQueryInfo.Create;
  try
      with TSql do begin
          Clear;
          AddSql(' Update TB_Master Set ');
          AddSql('   UpState = '''+UpState+''' ');
          AddSql(' Where ExamDate = '''+ExamDate+''' ');
          AddSql('   And ExamSeq = '''+ExamSeq+''' ');

          LocalExcute;
      end;
  finally
      TSql.Free;
  end;
end;

function TDM.SaveSpcId(Spcid, ExamDate, ExamSeq: string):boolean;
var
  TSql:TQueryInfo;
begin
  Result:= False;
  if Trim(Spcid) = '' then exit;

  TSql:= TQueryInfo.Create;

  try
      with TSql do begin
          Clear;
          AddSql(' Update TB_Master Set BarCode = '''+Spcid+''' ');
          AddSql(' Where ExamDate = '''+ExamDate+''' ');
          AddSql('   And ExamSeq  = '''+ExamSeq+''' ');
          Result:= LocalExcute;
      end;

  finally
      TSql.Free;
  end;

end;

function TDM.DeleteOneData(ExamDate, ExamSeq: string): boolean;
var
  TSql:TQueryInfo;
begin
  Result:= False;
  TSql:= TQueryInfo.Create;

  try
      with TSql do begin
          Clear;
          AddSql(' Delete From TB_Result ');
          AddSql(' Where ExamDate = '''+ExamDate+''' ');
          AddSql('   And ExamSeq  = '''+ExamSeq+''' ');
          LocalExcute;

          Clear;
          AddSql(' Delete From TB_Master ');
          AddSql(' Where ExamDate = '''+ExamDate+''' ');
          AddSql('   And ExamSeq  = '''+ExamSeq+''' ');
          Result:= LocalExcute;
      end;

  finally
      TSql.Free;
  end;

end;

function TDM.SaveOrder(TMaster:TIfMaster):string;
var
  TSql: TQueryInfo;
  QryEx:TADOQuery;
  UpCode:string;
begin
  Result:= '';

  UpCode:= TCode.GetUpCode(TMaster.FExamCode);
  if UpCode = '' then exit;

  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Self);

  try
      with Tsql do begin
          Clear;
          SqlCmd:= ' Select * FRom TB_Result '+
                   ' Where ExamDate = '''+TMaster.FExamDate+''' '+
                   '   And ExamSeq  = '''+TMaster.FExamSeq+'''  '+
                   '   And ExamCode = '''+TMaster.FExamCode+''' ';
          RCount:= LocalSelect(QryEx);

          if RCount = 0 then begin
              Clear;
              AddSql( ' Insert Into TB_Result    ');
              AddSql( ' ( ExamDate, ExamSeq, UpCode, OrdCode, ExamCode, ANO, OYN ) ');
              AddSql( ' Values                ');
              AddSql( ' (  '''+TMaster.FExamDate+'''   ');
              AddSql( '  , '''+TMaster.FExamSeq+'''    ');
              AddSql( '  , '''+UpCode+'''              ');
              AddSql( '  , '''+TMaster.FExamCode+'''   ');
              AddSql( '  , '''+TMaster.FExamCode+'''   ');
              AddSql( '  , '''+TMaster.FOrdSeq+'''     ');
              AddSql( '  , ''Y''     )');
          end
          else begin
              Clear;
              AddSql( ' Update TB_Result Set                         ');
              AddSql('     UpCode = '''+UpCode+'''                   ');
              AddSql('   , ANO = '''+TMaster.FOrdSeq+'''             ');
              AddSql('   , OYN = ''Y''           ');
              AddSql( ' Where ExamDate = '''+TMaster.FExamDate+'''   ');
              AddSql( '   And ExamSeq  = '''+TMaster.FExamSeq+'''    ');
              AddSql( '   And ExamCode   = '''+TMaster.FExamCode+''' ');
          end;

          if LocalExcute = True then
              Result:= UpCode;
      end;
  finally
      QryEx.Free;
      Tsql.Free;
  end;
end;

procedure TDM.SaveMaster(TMaster: TIfMaster);
var
  TSql:TQueryInfo;
  QryEx:TAdoQuery;
begin
  TSql:= TQueryInfo.Create;
  QryEx:= TADOQuery.Create(Application);

  try
      with TSql do begin
          Clear;
          AddSql(' Select * From TB_Master ');
          AddSql(' Where ExamDate = '''+TMaster.FExamDate+''' ');
          AddSql('   And ExamSeq  = '''+TMaster.FExamSeq+''' ');

          RCount:= LocalSelect(QryEx);

          if RCount = 0 then begin
              Clear;
              AddSql(' Insert Into TB_Master ');
              AddSql(' ( ExamDate, ExamSeq, ExamTime, BarCode, ReceiptNo, ODT, OrdCode, PId, PNM, RPOS, ADT, ANO, SLP, UpState ) ');
              AddSql(' Values ');
              AddSql(' ( '''+TMaster.FExamDate+''' ');
              AddSql(' , '''+TMaster.FExamSeq+'''  ');
              AddSql(' , '''+TMaster.FRcvTime+'''  ');
              AddSql(' , '''+TMaster.FBarCode+'''  ');
              AddSql(' , '''+TMaster.FANO+'''    ');
              AddSql(' , '''+TMaster.FOrdDate+'''  ');
              AddSql(' , '''+TMaster.FOrdCode+'''  ');
              AddSql(' , '''+TMaster.FPId+'''    ');
              AddSql(' , '''+TMaster.FPNm+'''    ');
              AddSql(' , '''+TMaster.FRack+'-'+TMaster.FPos+'''    ');
              AddSql(' , '''+TMaster.FADT+'''    ');
              AddSql(' , '''+TMaster.FANO+'''    ');
              AddSql(' , '''+TMaster.FSLP+'''    ');
              AddSql(' , '''+TMaster.FUpState+'''  )');
          end
          else begin
              Clear;
              AddSql(' Update TB_Master Set ');
              AddSql('     BarCode = '''+TMaster.FBarCode+''' ');
              AddSql('   , PId     = '''+TMaster.FPID+''' ');
              AddSql('   , PNM     = '''+TMaster.FPNM+''' ');
              AddSql('   , RPOS    = '''+TMaster.FRack+'-'+TMaster.FPos+''' ');
              AddSql('   , ADT     = '''+TMaster.FADT+''' ');
              AddSql('   , ANO     = '''+TMaster.FANO+''' ');
              AddSql('   , SLP     = '''+TMaster.FSLP+''' ');
              AddSql('   , UpState = '''+TMaster.FUpState+''' ');
              AddSql(' Where ExamDate = '''+TMaster.FExamDate+''' ');
              AddSql('   And ExamSeq  = '''+TMaster.FExamSeq+''' ');
          end;

          LocalExcute;
      end;
  finally
      QryEx.Free;
      TSql.Free;
  end;
end;

function TDM.GetExamUpCode(IfCode: string):Variant;
var
  TSql: TQueryInfo;
  QryEx: TAdoQuery;
begin
  Result:= VarArrayCreate([0,2], varVariant);
  Result[0]:='';
  Result[1]:='';
  Result[2]:='';

  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Application);

  try
      with TSql do begin
        Clear;
        SqlCmd:= ' Select ExamCode, UpCode, Abbr  From TB_CodeInfo '+
                 ' Where IfCode = '''+IfCode+''' ';
        RCount:= LocalSelect(QryEx);

        if RCount > 0 then begin
            Result[0]:= QryEx.FieldByName('ExamCode').AsString;
            Result[1]:= QryEx.FieldByName('UpCode').AsString;
            Result[2]:= QryEx.FieldByName('Abbr').AsString;
        end;
    end;

  finally
      TSql.Free;
      QryEx.Free;
  end;
end;

function TDM.CheckSetCode(ExamCode: string): boolean;
var
  TSql: TQueryInfo;
  QryEx: TAdoQuery;
begin
  Result:= False;

  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Application);

  try
      with TSql do begin
        Clear;
        SqlCmd:= ' Select *  From TB_CodeInfo Where ExamCode = '''+ExamCode+''' ';
        RCount:= LocalSelect(QryEx);

        if RCount > 0 then
            Result:= True;
      end;

  finally
      TSql.Free;
      QryEx.Free;
  end;

end;

function TDM.UploadHost_One_YJ(ReceiptNo, OCD, ECD,ResVal:string):boolean;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:= 'Local 테스트중입니다!';
      exit;
  end;

  with qrySUp do begin
      Close;
      SQL.Text:=' update SLA_LABRESULT set Result = '''+ResVal+'''       ' +#13#10+
                '                        , ResultDate = To_Char(SysDate, ''YYYY-MM-DD'') '+#13#10+
                '                        , ResultTime = To_Char(SysDate, ''HH24:MI:SS'')  '+#13#10+
                ' where ReceiptNo = '+ReceiptNo+' '+#13#10+
                '   And ORDERCODE = '''+OCD+'''   '+#13#10+
                '   And LABCODE   = '''+ECD+'''   '+#13#10+
                '   And transflag < ''2'' ';
      if F_Main.Debug1.Checked then
          SQL.SaveToFile('결과업로드.txt');

      try
          ExecSQL;
          Result:= True;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= 'SLA_LABRESULT 결과전송 에러 입니다! 에러메세지->'+ e.Message;
          end;
      end;
  end;

end;



function TDM.getOrdCode(PanelName: string): string;
var
  TSql: TQueryInfo;
  QryEx: TAdoQuery;
begin
  Result:= '';

  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Application);

  try
      with TSql do begin
        Clear;
        SqlCmd:= ' Select ORDCODE From TB_CodeInfo '+
                 ' Where PANEL = '''+PanelName+''' ';
        RCount:= LocalSelect(QryEx);

        if RCount > 0 then begin
            Result:= QryEx.Fields[0].AsString;
        end;
    end;

  finally
      TSql.Free;
      QryEx.Free;
  end;

end;

function TDM.GetExamCode(OrdCode, IfCode: string): string;
var
  TSql: TQueryInfo;
  QryEx: TAdoQuery;
begin
  Result:= '';

  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Application);

  try
      with TSql do begin
        Clear;
        SqlCmd:= ' Select ExamCODE From TB_CodeInfo '+
                 ' Where ORDCODE = '''+OrdCode+'''  '+
                 '   And UpCode= '''+IfCode+''' ';

        RCount:= LocalSelect(QryEx);

        if RCount > 0 then begin
            Result:= QryEx.Fields[0].AsString;
        end;
    end;

  finally
      TSql.Free;
      QryEx.Free;
  end;

end;

function TDM.DownLoadOrder_YJ(TMaster:TIfMaster): boolean;
var
  OCD:string;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:=  'Local TEST중입니다.';
      exit;
  end;

  with qrySOrder do begin
      Close;
      SQL.Text:=' Select DISTINCT ReceiptNo, IpdOpd, ReceiptDate, ORDERCODE, PTno, SName, Age, '+
                '        SUBSTR(ReceiptTime, 1, 5) ReceiptTime, Sex, BI, DeptCode, '+
                '        WardCode, Roomcode, BillFlag DrCode, JStatus,  SPECIMENNUM '+
                ' From SLA_LabMaster '+
                ' Where SPECIMENNUM = '+TMaster.FBarCode+
                //'   And ORDERCODE = '''+TClass.FOrdCode+''' '+
                '   And JsTATUS < ''2'' ';
      try
          Open;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= 'SLA_LABMASTER 오더체크 에러 입니다! 에러메세지->'+ e.Message;
              exit;
          end;
      end;

      if RecordCount > 0 then begin
          while Not Eof do begin
              OCD:= FieldByName('ORDERCODE').AsString;
              if CheckOrderCode(OCD) then begin
                  TMaster.FOrdCode := OCD;
                  TMaster.FPID  := FieldByName('PTNO').AsString;
                  TMaster.FPNM  := FieldByName('SName').AsString;
                  TMaster.FAge  := FieldByName('AGE').AsString;
                  TMaster.FSex  := FieldByName('SEX').AsString;
                  TMaster.FDept := FieldByName('DeptCode').AsString;
                  TMaster.ReceiptNo:= FieldByName('ReceiptNo').AsInteger;
                  Result:= True;
                  TMaster.FOrdState:= 'Y';
                  Exit;
              end;
              Next;
          end;
      end;
  end;
end;

function TDM.UploadState_YJ(ReceiptNo, OrdCode, UpState: string): boolean;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:=  'Local TEST중입니다.';
      exit;
  end;


  with qrySUp1 do begin
      Close;
      SQL.Text:= ' Update SLA_LABMASTER SET JSTATUS = '''+UpState+''' '+
                 ' Where ReceiptNo = '+ReceiptNo+
                 '   And OrderCode = '''+OrdCode+''' '+
                 '   And JStatus < ''3'' ';
      try
          ExecSQL;
          Result:= True;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= 'SLA_LABMASTER JSTATUS변경 에러입니다! 에러메세지->'+ e.Message;
          end;
      end;
  end;

end;

function TDM.GetExamPanelCode(PCode, IfCode: string): string;
var
  TSql: TQueryInfo;
  QryEx: TAdoQuery;
begin
  Result:= '';

  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Application);

  try
      with TSql do begin
        Clear;
        SqlCmd:= ' Select ExamCODE From TB_CodeInfo '+
                 ' Where PANEL = '''+PCode+'''  '+
                 '   And UpCode= '''+IfCode+''' ';

        RCount:= LocalSelect(QryEx);

        if RCount > 0 then begin
            Result:= QryEx.Fields[0].AsString;
        end;
    end;

  finally
      TSql.Free;
      QryEx.Free;
  end;

end;

function TDM.CheckOrderCode(OrdCode: string): boolean;
var
  TSql: TQueryInfo;
  QryEx: TAdoQuery;
begin
  Result:= False;

  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Application);

  try
      with TSql do begin
        Clear;
        SqlCmd:= ' Select ExamCODE From TB_CodeInfo '+
                 ' Where OrdCode = '''+OrdCode+'''  ';

        RCount:= LocalSelect(QryEx);

        if RCount > 0 then begin
            Result:= True;
        end;
    end;

  finally
      TSql.Free;
      QryEx.Free;
  end;
end;

function TDM.DownLoadOrder_PTM(TMaster: TIfMaster): boolean;
var
  ECD:string;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:=  'Local TEST중입니다.';
      exit;
  end;

  with qrySOrder do begin
      Close;
      SQL.Text:= ' select a.bunho, a.suname, a.sex, a.age, b.order_date, c.gumsa_code '+
                 ' from cpl03d1 c, cpl03m1 b, cpl0201 a                    '+
                 ' where a.specimen_ser = '''+TMaster.FBarCode+'''         '+ //<-검체번호(바코드)
                 '   and b.fkcpl0201    = a.pkcpl0201                     '+
                 '   and c.fkcpl0201    = b.fkcpl0201                     '+
                 '   and c.source_gumsa = b.gumsa_code                    '+
                 '   and c.cpl_result is null                             ';
      try
          Open;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= '오더체크 에러 입니다! 에러메세지->'+ e.Message+' SQL=>'+
              SQL.Text;
              exit;
          end;
      end;

      if RecordCount > 0 then begin
          while Not Eof do begin
              ECD:= FieldByName('gumsa_code').AsString;
              if CheckSetCode(ECD) then begin
                  TMaster.FPID  := FieldByName('bunho').AsString;
                  TMaster.FPNM  := FieldByName('suname').AsString;
                  TMaster.FAge  := FieldByName('age').AsString;
                  TMaster.FSex  := FieldByName('sex').AsString;
                  TMaster.FOrdDate := FieldByName('order_date').AsString;
                  Result:= True;
                  TMaster.FOrdState:= 'Y';
                  Exit;
              end;
              Next;
          end;
      end;
  end;
end;

function TDM.UploadHost_One_PTM(BCD, ECD, ResVal: string): boolean;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:= 'Local 테스트중입니다!';
      exit;
  end;

  with qrySUp do begin
      Close;
      SQL.Text:= ' Insert Into CPL_JANGBI_CARDIAC (SYS_DATE, SPECIMEN_SER, GUMSA_CODE, CPL_RESULT, '+
                 ' RESULT_DATE, SEND_YN ) values ( sysdate, '''+BCD+''', '''+Ecd+''', '''+ResVal+''', sysdate, null ) ';

      if F_Main.Debug1.Checked then
          SQL.SaveToFile('결과업로드.txt');

      try
          ExecSQL;
          Result:= True;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= 'CPL_JANGBI_CARDIAC 결과전송 에러 입니다! 에러메세지->'+ e.Message+' SQL=>'+
              SQL.Text;
          end;
      end;
  end;

end;

function TDM.DownLoadOrder_DSW(TMaster: TIfMaster): boolean;
var
  ECD:string;
  vAgeSex:Variant;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:=  'Local TEST중입니다.';
      exit;
  end;

  with qrySOrder do begin
      Close;
      SQL.Text:=' select R.ACPTDATE                '+#13#10+
                '      , R.ACPTSEQ                 '+#13#10+
                '      , R.RESULTITEMCODE          '+#13#10+
                '      , R.RESULTCODE              '+#13#10+
                '      , R.SUBRESULTCODE           '+#13#10+
                '      , R.ORDERCODE               '+#13#10+
                '      , R.INSTCODE                '+#13#10+
                '      , R.PATNO                   '+#13#10+
                '      , R.RESULTSTATE             '+#13#10+
                '      , R.RESULT                  '+#13#10+
                '      , W.CLASS                   '+#13#10+
                '      , W.SLIPCODE                '+#13#10+
                '      , W.DEPTCODE                '+#13#10+
                '      , W.ORDERNAME               '+#13#10+
                '      , P.PATNAME                 '+#13#10+
                '      , P.JUMIN1                  '+#13#10+
                '      , P.JUMIN2                  '+#13#10+
                ' from                             '+#13#10+
                '   MH_LABRESULT     R             '+#13#10+
                ' , MH_LABREGISTINFO W             '+#13#10+
                ' , MH_PATINFO       P             '+#13#10+
                ' where R.ACPTDATE = W.ACPTDATE    '+#13#10+
                '   and R.ACPTSEQ  = W.ACPTSEQ     '+#13#10+
                '   and R.PATNO    = W.PATNO       '+#13#10+
                '   and R.PATNO    = P.PATNO       '+#13#10+
                '   and R.SPCMNUM  = '''+TMaster.FBarCode+'''  '+
                '   and R.RESULTSTATE in (0,1,2) ';  
      try
          Open;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= '오더체크 에러 입니다! 에러메세지->'+ e.Message+' SQL=>'+
              SQL.Text;
              exit;
          end;
      end;

      if RecordCount > 0 then begin
          while Not Eof do begin
              ECD:= FieldByName('ORDERCODE').AsString;
              if CheckSetCode(ECD) then begin
                  TMaster.FPID := FieldByName('PATNO').AsString;
                  TMaster.FPNM := FieldByName('PATNAME').AsString;
                  vAgeSex:= GetAgeSex(FieldByName('JUMIN1').AsString+FieldByName('JUMIN2').AsString);
                  TMaster.FAge := vAgeSex[0];
                  TMaster.FSex := vAgeSex[1];
                  TMaster.FADT := FormatDateTime('yyyymmdd', FieldByName('ACPTDATE').AsDateTime);
                  TMaster.FANO := IntToStr(FieldByName('ACPTSEQ').AsInteger);
                  Result:= True;
                  TMaster.FOrdState:= 'Y';
                  Exit;
              end;
              Next;
          end;
      end;
  end;
end;

function TDM.UploadHost_One_DSW(BCD, ECD, ResVal, LH: string; var ErrMsg:string): boolean;
begin
  Result:= False;
  ErrMsg:= '';

  if SvrConnection = False then begin
      TGlobal.ErrMsg:= 'Local 테스트중입니다!';
      ErrMsg:= 'Local 테스트중입니다!';
      exit;
  end;

  with qrySUp do begin
      Close;
      SQL.Text:=' Update MH_LABRESULT Set           '+#13#10+
                '     RESULTSTATE  = 1              '+#13#10+
                '   , RESULT       = '''+ResVal+''' '+#13#10+
                '   , RESULTINDATE = sysdate        '+#13#10+
                '   , HIGHLOWFLAG  = '''+LH+'''     '+#13#10+
                ' Where SPCMNUM   = '''+BCD+'''     '+#13#10+
                '   And ORDERCODE = '''+ECD+'''     '+#13#10+
                '   And RESULTSTATE in (0,1,2)      ';

      if F_Main.Debug1.Checked then
          SQL.SaveToFile('결과업로드.txt');

      try
          ExecSQL;
          Result:= True;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= 'MH_LABRESULT 결과전송 에러 입니다! 에러메세지->'+ e.Message+' SQL=>'+
              SQL.Text;

              ErrMsg:=TGlobal.ErrMsg;
          end;
      end;
  end;

end;

function TDM.UploadHost_State_DSW(ADT, ANO, PID, STA: string): boolean;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:= 'Local 테스트중입니다!';
      exit;
  end;

  with qrySUp do begin
      Close;
      SQL.Text:=' Update MH_LABREGISTINFO Set                         '+#13#10+
                '   RESULTDATE  = sysdate                             '+#13#10+
                ' , RESULTSTATE = 1                                   '+#13#10+
                ' , RSVACTFLAG  = 5                                   '+#13#10+
                ' Where ACPTDATE = to_date('''+ADT+''', ''yyyymmdd'') '+#13#10+
                '   And ACPTSEQ  = '+ANO+'                            '+#13#10+
                '   And PATNO    = '''+PID+'''                        '+#13#10+
                '   And RSVACTFLAG = 4                                '+#13#10+
                '   And RESULTSTATE in (0,1,2) ';

      if F_Main.Debug1.Checked then
          SQL.SaveToFile('STATE변경.txt');

      try
          ExecSQL;
          Result:= True;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= 'MH_LABREGISTINFO STATE변경 에러 입니다! 에러메세지->'+ e.Message+' SQL=>'+
              SQL.Text;
          end;
      end;
  end;

end;

function TDM.GetANO(BCD, ECD: string): string;
begin
  Result:= '';

  if SvrConnection = False then begin
      TGlobal.ErrMsg:= 'Local 테스트중입니다!';
      exit;
  end;

  with qrySUp do begin
      Close;
      SQL.Text:=' Select ACPTSEQ From MH_LABRESULT '+#13#10+
                ' Where SPCMNUM   = '''+BCD+'''     '+#13#10+
                '   And ORDERCODE = '''+ECD+'''     ';
      try
          Open;
          if RecordCount > 0 then
              Result:= Fields[0].AsString;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= 'MH_LABRESULT 결과전송 에러 입니다! 에러메세지->'+ e.Message+' SQL=>'+
              SQL.Text;
          end;
      end;
  end;

end;

function TDM.CheckSetCode_ORD(OrdCode: string): boolean;
var
  TSql: TQueryInfo;
  QryEx: TAdoQuery;
begin
  Result:= False;

  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Application);

  try
      with TSql do begin
        Clear;
        SqlCmd:= ' Select *  From TB_CodeInfo Where OrdCode = '''+OrdCode+''' ';
        RCount:= LocalSelect(QryEx);

        if RCount > 0 then
            Result:= True;
      end;

  finally
      TSql.Free;
      QryEx.Free;
  end;

end;

function TDM.DownLoadOrder_KY(TMaster: TIfMaster): boolean;
var
  ECD:string;
  vAgeSex:Variant;
  ADT, SLP, LNO:string;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:=  'Local TEST중입니다.';
      exit;
  end;

  if DownLoadOrder_KY_BCD(TMaster) then
  begin
      Result:= True;
      exit;
  end;

  if Not GetJusuData(TMaster.FBarCode, ADT, SLP, LNO) then
      exit;

  with qrySOrder do begin
      Close;
      SQL.Text:=' select j.jeobsudt                         '+#13#10+
                '      , j.slipno1                          '+#13#10+
                '      , j.slipno2                          '+#13#10+
                '      , j.ptno                             '+#13#10+
                '      , j.deptcode                         '+#13#10+
                '      , j.status                           '+#13#10+
                '      , r.itemcd                           '+#13#10+
                '      , r.result1                          '+#13#10+
                '      , r.verify                           '+#13#10+
                '      , p.sname, p.jumin1||p.jumin2 as jno '+#13#10+
                ' from twexam_general j                     '+#13#10+
                '    , twexam_general_sub r                 '+#13#10+
                '    , twbas_patient p                      '+#13#10+
                ' where j.jeobsudt = r.jeobsudt             '+#13#10+
                '   and j.slipno1 = r.slipno1               '+#13#10+
                '   and j.slipno2 = r.slipno2               '+#13#10+
                '   and j.ptno = p.ptno                     '+#13#10+
                '   and j.jeobsudt = to_date('''+ADT+''', ''yyyymmdd'') '+#13#10+
                '   and j.slipno1 = '''+SLP+'''             '+#13#10+
                '   and j.slipno2 = '''+LNO+'''             ';
                //'   and r.verify <> ''Y''                   ';
      try
          if SvrTEST then
              sql.SaveToFile('오더테스트2.sql');

          Open;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= '오더체크 에러 입니다! 에러메세지->'+ e.Message+' SQL=>'+
              SQL.Text;
              exit;
          end;
      end;

      if RecordCount > 0 then begin
          while Not Eof do begin
              ECD:= FieldByName('itemcd').AsString;
              if CheckSetCode(ECD) then begin
                  TMaster.FPID := FieldByName('ptno').AsString;
                  TMaster.FPNM := FieldByName('sname').AsString;
                  vAgeSex:= GetAgeSex(FieldByName('jno').AsString);
                  TMaster.FAge := vAgeSex[0];
                  TMaster.FSex := vAgeSex[1];
                  TMaster.FADT := ADT;
                  TMaster.FSLP := SLP;
                  TMaster.FANO := LNO;
                  Result:= True;
                  TMaster.FOrdState:= 'Y';
                  Exit;
              end;
              Next;
          end;
      end;
  end;
end;

function TDM.DownLoadOrder_KY_BCD(TMaster: TIfMaster): boolean;
var
  ECD:string;
  vAgeSex:Variant;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:=  'Local TEST중입니다.';
      exit;
  end;

  with qrySOrder do begin
      Close;
      SQL.Text:=' select j.jeobsudt                         '+#13#10+
                '      , j.slipno1                          '+#13#10+
                '      , j.slipno2                          '+#13#10+
                '      , j.ptno                             '+#13#10+
                '      , j.deptcode                         '+#13#10+
                '      , j.status                           '+#13#10+
                '      , r.itemcd                           '+#13#10+
                '      , r.result1                          '+#13#10+
                '      , r.verify                           '+#13#10+
                '      , p.sname, p.jumin1||p.jumin2 as jno '+#13#10+
                ' from twexam_general j                     '+#13#10+
                '    , twexam_general_sub r                 '+#13#10+
                '    , twbas_patient p                      '+#13#10+
                ' where j.jeobsudt = r.jeobsudt             '+#13#10+
                '   and j.slipno1 = r.slipno1               '+#13#10+
                '   and j.slipno2 = r.slipno2               '+#13#10+
                '   and j.ptno = p.ptno                     '+#13#10+
                '   and j.barcode = '''+TMaster.FBarCode+''' '+#13#10;
                //'   and r.verify <> ''Y''                   ';
      try
          if SvrTEST then
              sql.SaveToFile('오더테스트1.sql');

          Open;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= '오더체크 에러 입니다! 에러메세지->'+ e.Message+' SQL=>'+
              SQL.Text;
              exit;
          end;
      end;

      if RecordCount > 0 then begin
          while Not Eof do begin
              ECD:= FieldByName('itemcd').AsString;
              if CheckSetCode(ECD) then begin
                  TMaster.FPID := FieldByName('ptno').AsString;
                  TMaster.FPNM := FieldByName('sname').AsString;
                  vAgeSex:= GetAgeSex(FieldByName('jno').AsString);
                  TMaster.FAge := vAgeSex[0];
                  TMaster.FSex := vAgeSex[1];
                  TMaster.FADT := FormatDateTime('yyyymmdd', FieldByName('jeobsudt').AsDateTime);
                  TMaster.FSLP := FieldByName('slipno1').AsString;
                  TMaster.FANO := FieldByName('slipno2').AsString;
                  TMaster.FDept:= FieldByName('deptcode').AsString;
                  Result:= True;
                  TMaster.FOrdState:= 'Y';
                  Exit;
              end;
              Next;
          end;
      end;
  end;
end;

function TDM.UploadHost_One_KY(ADT, SLP, ANO, ECD, PID, RES:string;
  var ErrMsg: string): boolean;
begin
  Result:= False;
  ErrMsg:= '';

  if SvrConnection = False then begin
      TGlobal.ErrMsg:= 'Local 테스트중입니다!';
      ErrMsg:= 'Local 테스트중입니다!';
      exit;
  end;

  with qrySUp do begin
      Close;
      SQL.Text:=' Update twexam_general_sub Set           '+#13#10+
                '     result1      = '''+RES+'''          '+#13#10+
                '   , outclassno   = ''34''               '+#13#10+  //장비코드
                ' where jeobsudt = to_date('''+ADT+''', ''yyyymmdd'') '+#13#10+
                '   and slipno1 = '''+SLP+'''             '+#13#10+
                '   and slipno2 = '''+ANO+'''             '+#13#10+
                '   and itemcd  = '''+ECD+'''             '+#13#10+
                '   and ptno    = '''+PID+'''             '+#13#10+
                '   and verify <> ''Y''                   ';
      try
          if SvrTEST then
              sql.SaveToFile('결과업로드.sql');

          ExecSQL;
          Result:= True;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= 'twexam_general_sub 결과전송 에러 입니다! 에러메세지->'+ e.Message+' SQL=>'+
              SQL.Text;

              ErrMsg:=TGlobal.ErrMsg;
          end;
      end;
  end;
end;

function TDM.UploadHost_State_KY(ADT, SLP, ANO, PID, STA: string): boolean;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:= 'Local 테스트중입니다!';
      exit;
  end;

  with qrySUp do begin
      Close;
      SQL.Text:=' Update twexam_general Set               '+#13#10+
                '     status  = '''+STA+'''               '+#13#10+
                ' where jeobsudt = to_date('''+ADT+''', ''yyyymmdd'') '+#13#10+
                '   and slipno1 = '''+SLP+'''             '+#13#10+
                '   and slipno2 = '''+ANO+'''             '+#13#10+
                '   and ptno    = '''+PID+'''             '+#13#10+
                '   and status  in (''U'', ''R'', ''P'')' ;
      try
          if SvrTEST then
              sql.SaveToFile('status업로드.sql');

          ExecSQL;
          Result:= True;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= 'twexam_general 상태변경 에러 입니다! 에러메세지->'+ e.Message+' SQL=>'+
              SQL.Text;
          end;
      end;
  end;
end;



function TDM.GetExamCode_Var(IfCode: string): Variant;
var
  TSql: TQueryInfo;
  QryEx: TAdoQuery;
  i:integer;
  VCode:Variant;
  Tmp:Variant;
begin
  Tmp:= VarArrayCreate([0,0], varVariant);

  Result:= Tmp;

  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Application);

  try
      with TSql do begin
        Clear;
        SqlCmd:= ' Select ExamCODE From TB_CodeInfo '+
                 ' Where UpCode= '''+IfCode+''' ';

        RCount:= LocalSelect(QryEx);

        if RCount > 0 then begin
            i:=-1;
            VCode:= VarArrayCreate([0, RCount-1], varVariant);

            while Not QryEx.Eof do begin
                Inc(i);
                VCode[i]:= QryEx.Fields[0].AsString;
                QryEx.Next;
            end;
        end
        else begin
            VCode:= VarArrayCreate([0,0], varVariant);
            VCode[0]:='';
        end;
    end;

    Result:= VCode;

  finally
      TSql.Free;
      QryEx.Free;
  end;

end;

function TDM.UploadResult(TMaster:TIfMaster): boolean;
var
  UpCnt:integer;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:= 'Local 테스트중입니다!';
      exit;
  end;

  UpCnt:=0;

  with qryUp1 do begin
      Close;
      SQL.Text:= ' SELECT  M.BarCode                         '+#13#10+
                 '       , R.ExamCode                        '+#13#10+
                 '       , R.Result                          '+#13#10+
                 '       , R.FLAG                            '+#13#10+
                 '       , R.UpCode                          '+#13#10+
                 ' FROM TB_Master AS M, TB_Result AS R       '+#13#10+
                 ' where M.ExamDate = '''+TMaster.FExamDate+'''       '+#13#10+
                 '   And M.ExamSeq  = '''+TMaster.FExamSeq+'''        '+#13#10+
                 '   And M.ExamDate = R.ExamDate             '+#13#10+
                 '   And M.ExamSeq  = R.ExamSeq              ';
      Open;

      if RecordCount > 0 then begin
          while Not Eof do begin
              TMaster.FUpCode:= FieldByName('UpCode').AsString;
              TMaster.FResult := FieldByName('Result').AsString;
              TMaster.FFlag := FieldByName('FLAG').AsString;

              SetDownCode(TMaster);

              if TMaster.IsDownCodeOK = True then
                  if UploadHosp_One_DJI(TMaster) then
                      Inc(UpCnt);

              Next;
          end;
      end;

      if UpCnt > 0 then
          Result:= True;
          //Result:= UploadState_DJ(TMaster);
  end;
end;

function TDM.DownLoadPAT_JAIN(TMaster: TIfMaster): boolean;
var
  ECD:string;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:=  'Local TEST중입니다.';
      exit;
  end;

  with qrySOrder do begin
      Close;
      SQL.Text:=' Select SCP41PCODE                  '+#13#10+
                '      , SCP41IDNOA                  '+#13#10+
                '      , SCP41NAME                   '+#13#10+
                '      , SCP41SNDYN                  '+#13#10+
                ' From JAIN_SCP.SCPRST41             '+#13#10+
                ' Where SCP41SPMNO2 = '''+TMaster.FBarCode+'''  ';
      try
          if SvrTEST then
              SQL.SaveToFile('환자조회.sql');

          Open;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= '환자정보 체크 에러 입니다! 에러메세지->'+ e.Message+' SQL=>'+
              SQL.Text;
              ShowMessage(e.Message);
          end;
      end;

      if RecordCount > 0 then begin
            TMaster.FPID:= FieldByName('SCP41IDNOA').AsString;
            TMaster.FPNM:= FieldByName('SCP41NAME').AsString;
            //TMaster.FSex:= FieldByName('IDSEX').AsString;
            if FieldByName('SCP41SNDYN').AsString <> 'Y' then
            begin
                TMaster.FOrdState:= 'Y';
                Result:= True;
                exit;
            end;
      end;
  end;
end;

function TDM.UploadHost_One_JAIN(BCD, ECD, ResVal: string): boolean;
var
  sNow:string;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:= 'Local 테스트중입니다!';
      exit;
  end;

  sNow:= FormatDateTime('yyyymmdd', now);

  with qrySUp do begin
      Close;
      SQL.Text:= ' SELECT SCP41SNDYN                '+#13#10+
                 ' From JAIN_SCP.SCPRST41           '+#13#10+
                 ' Where SCP41SPMNO2 = '''+BCD+'''  '+#13#10+
                 '   and SCP41SNDYN = ''Y''         ';
      try
          if SvrTEST then
              SQL.SaveToFile('업데이트1.sql');

          Open;

      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= '업데이트_환자조회 에러 입니다! 에러메세지->'+ e.Message+' SQL=>'+SQL.Text;
              ShowMessage(e.Message);
          end;
      end;

      if RecordCount > 0 then exit;

      SQL.Text:=' UPDATE JAIN_SCP.SCPRST42 SET           '+#13#10+
                '     SCP42TSTDAT = '''+sNow+'''         '+#13#10+
                '  ,  SCP42RSTCD  = ''N''                '+#13#10+   //결과형식 => 숫자 : 'N', 문자 : 'X', 장문 : 'R'
                '  ,  SCP42RESULT = '''+ResVal+'''       '+#13#10+
                ' Where SCP42SPMNO2 = '''+BCD+'''        '+#13#10+
                '   And SCP42SUGACD = '''+ECD+'''        ';

                {'   And SCP42SUGACD IN ( SELECT SCP52SUGACD FROM SCPMST52 '+
                '                        WHERE SCP52PCODE = ''10''        '+
                '                          AND SCP52TMCH  = '''+TGlobal.FICode+''') ' ;   }

      try
          if SvrTEST then
              SQL.SaveToFile('결과등록.sql');

          ExecSQL;
          Result:= True;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= '결과전송 에러 입니다! 에러메세지->'+ e.Message+' SQL=>'+
              SQL.Text;
              ShowMessage(e.Message);
          end;
      end;
  end;


end;

function TDM.UploadHost_State_JAIN(BCD, STA: string): boolean;
var
  sNow:string;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:= 'Local 테스트중입니다!';
      exit;
  end;

  sNow:= FormatDateTime('yyyymmdd', now);

  with qrySUp do begin
      Close;
      SQL.Text:= ' UPDATE JAIN_SCP.SCPRST41 SET       '+#13#10+
                 '     SCP41TSTDAT   = '''+sNow+'''   '+#13#10+
                 '   , SCP41SNDYN    = ''N''          '+#13#10+
                 '   , SCP41RSTYN    = ''Y''          '+#13#10+
                 '   , SCP41TSTUID   = '''+TGlobal.FUserID+''' '+#13#10+
                 ' Where SCP41SPMNO2 = '''+BCD+'''    '+#13#10+
                 '   And SCP41SNDYN  <> ''Y''  ';

      try
          if SvrTEST then
              SQL.SaveToFile('상태변경.sql');

          ExecSQL;
          Result:= True;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= '상태변경 에러 입니다! 에러메세지->'+ e.Message+' SQL=>'+
              SQL.Text;
              ShowMessage(e.Message);
          end;
      end;
  end;

end;

function TDM.DownLoadOrder_JAIN(TMaster: TIfMaster): boolean;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:= 'Local 테스트중입니다!';
      exit;
  end;

  with qrySOrder do begin
      Close;
      SQL.Text:=' Select SCP42SUGACD                            '+#13#10+
                ' From JAIN_SCP.SCPRST42                        '+#13#10+
                ' Where SCP42SPMNO2 = '''+TMaster.FBarCode+'''  ';
                //'   And SCP42SUGACD = '''+TMaster.FExamCode+''' ';
      try
          if SvrTEST then
              SQL.SaveToFile('오더조회.sql');

          Open;

          while Not Eof do begin
              if TCode.SetCode_ECode(Fields[0].AsString) then
              begin
                  Result:= True;
                  TMaster.FOrdState:= 'Y';
                  exit;
              end;
              Next;
          end;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= '오더조회 에러 입니다! 에러메세지->'+ e.Message+' SQL=>'+ SQL.Text;
              ShowMessage(e.Message);
          end;
      end;
  end;
end;



function TDM.DownLoadOrder_JSD(TMaster: TIfMaster): boolean;
var
  ECD:string;
  vAgeSex:Variant;
  ADT, SLP, LNO:string;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:=  'Local TEST중입니다.';
      exit;
  end;

  with qrySOrder do begin
      Close;
      SQL.Text:=' SELECT A.HANGMOG_CODE                   '+#13#10+        
                '      , A.SPECIMEN_CODE                  '+#13#10+
                '      , A.EMERGENCY                      '+#13#10+
                '      , A.SPECIMEN_SER                   '+#13#10+
                '      , A.SOURCE_GUMSA                   '+#13#10+
                '      , A.LAB_NO                         '+#13#10+
                '      , A.CONFIRM_YN                     '+#13#10+
                '      , C.ORDER_DATE                     '+#13#10+
                '      , C.BUNHO                          '+#13#10+
                '      , C.SUNAME                         '+#13#10+
                '      , C.AGE                            '+#13#10+
                '      , C.SEX                            '+#13#10+
                ' FROM MEDI.CPL3020 A                     '+#13#10+
                '    , MEDI.CPL2010 C                     '+#13#10+
                ' WHERE A.SPECIMEN_SER = '''+TMaster.FBarCode+'''     '+#13#10+
                '   AND A.SPECIMEN_SER = C.SPECIMEN_SER   '+#13#10+
                '   AND A.HANGMOG_CODE = C.HANGMOG_CODE   '+#13#10+
                '   AND NVL(A.CONFIRM_YN, ''N'') = ''N''  ';
      try
          if SvrTEST then
              sql.SaveToFile('오더조회.sql');

          Open;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= '오더조회 에러 입니다! 에러메세지->'+ e.Message+' SQL=>'+ SQL.Text;
              ShowMessage(e.Message);
              exit;
          end;
      end;

      if RecordCount > 0 then begin
          while Not Eof do begin
              ECD:= FieldByName('HANGMOG_CODE').AsString;
              if TCode.SetCode_ECode(ECD) then begin
                  TMaster.FPID := FieldByName('BUNHO').AsString;
                  TMaster.FPNM := FieldByName('SUNAME').AsString;
                  TMaster.FAge := Trim(FieldByName('AGE').AsString);
                  TMaster.FSex := Trim(FieldByName('SEX').AsString);
                  TMaster.FANO := FieldByName('LAB_NO').AsString;
                  Result:= True;
                  TMaster.FOrdState:= 'Y';
                  Exit;
              end;
              Next;
          end;
      end;
  end;

end;

function TDM.UploadHost_One_JSD(BCD, ECD, IfCd, RES:string): boolean;
var
  spIn:TADOStoredProc;
begin
  Result:= False;
{
  with qrySUp do begin
      Close;
      SQL.Text:= 'exec medi.pr_cpl_insert_cpl0891( ''CHORUS'', '+
                 '  '''+TGlobal.FICode+''', '''+BCD+''', '''+IfCd+''', '+
                 '  '''+RES+''', '''+FormatDateTime('yyyymmdd',now)+''', '''' )';
      if SvrTEST then
          SQL.SaveToFile('.\프로시져실행!.sql');
      try
          ExecSQL;
          Result:= True;
      except
          on e:exception do begin
              ShowMessage(e.Message);
              exit;
          end;
      end;
  end; }

  spIn:= TADOStoredProc.Create(nil);
  try
      spIn.Connection:= TConnection.hospCon;

      with spIn do begin
          Close;
          ProcedureName:= 'MEDI.PR_CPL_INSERT_CPL0891';
          Parameters.CreateParameter('I_USER_ID', ftString, pdInput, 4000, '');
          Parameters.CreateParameter('I_JANGBI_CODE', ftString, pdInput, 4000, '');
          Parameters.CreateParameter('I_SPECIMEN_SER', ftString, pdInput, 4000, '');
          Parameters.CreateParameter('I_JANGBI_OUT_CODE', ftString, pdInput, 4000, '');
          Parameters.CreateParameter('I_CPL_RESULT', ftString, pdInput, 4000, '');
          Parameters.CreateParameter('I_RESULT_DATE', ftString, pdInput, 4000, '');
          Parameters.CreateParameter('I_RESULT_SEQ', ftString, pdInput, 4000, '');

          Parameters.ParamByName('I_USER_ID').Value:= 'CHORUS';
          Parameters.ParamByName('I_JANGBI_CODE').Value:= TGlobal.FICode;
          Parameters.ParamByName('I_SPECIMEN_SER').Value:= BCD;
          Parameters.ParamByName('I_JANGBI_OUT_CODE').Value:= IfCd;
          Parameters.ParamByName('I_CPL_RESULT').Value:= RES;
          Parameters.ParamByName('I_RESULT_DATE').Value:= FormatDateTime('yyyymmdd', now);
          Parameters.ParamByName('I_RESULT_SEQ').Value:= '';

          try
              ExecProc;
              Result:= True;
          except
              on e:exception do begin
                  ShowMessage(e.Message);
                  exit;
              end;
          end;
      end;
  finally
      spIn.Free;
  end;

end;

function TDM.DownLoadOrder_JEIL(TMaster: TIfMaster): boolean;
var
  HBCD, ECD:string;
begin
  Result:= False;

  if SvrConnection = false then begin
      ShowMessage('로컬테스트중 입니다!');
      exit;
  end;

  if TMaster.FQCYN = 'Y' then begin
      Result:= DownLoadOrder_QC_JEIL(TMaster);
      exit;
  end;

  HBCD:= MakeBarCode(TMaster.FBarCode);

  with qrySOrder do begin
      Close;
      SQL.Text:=' Select R.PRSNMRNO   PID         '+#13#10+   //Key  고객번호
                '      , R.PRSNVSDP   DPT         '+#13#10+   //Key  내원과/재원과
                '      , R.PRSNVSDT   IDT         '+#13#10+   //Key  내원일/입원일
                '      , R.PRSNLDTE   ODT         '+#13#10+   //Key  처방일자
                '      , R.PRSNSLIP   SLP         '+#13#10+   //Key  Slip코드
                '      , R.PRSNCODE   ECD         '+#13#10+   //Key  검사코드
                '      , R.PRSNSUBC   SUB         '+#13#10+   //Key  검사상세코드
                '      , R.PRSNORNO   ONO         '+#13#10+   //Key  ORDER NUMBER
                '      , R.PRSNORSQ   OSEQ        '+#13#10+   //Key  ORDER SEQUENCE
                '      , R.PRSNLABR  	LAB         '+#13#10+   //검사과
                '      , R.PRSNLBNO   BCD         '+#13#10+   //Lab Number
                '      , R.PRSNSPEC               '+#13#10+   //검체
                '      , R.PRSNCASE               '+#13#10+   //내원번호->장비코드로 사용
                '      , R.PRSNRSLT               '+#13#10+   //검사결과
                '      , R.PRSNCPMV               '+#13#10+   //보조결과
                '      , R.PRSNPWFL               '+#13#10+   //W/S 출력FLAG
                '      , R.PRSNRLFL               '+#13#10+   //검사PANIC FLAG
                '      , R.PRSNRMKF               '+#13#10+   //REMARKS FLAG
                '      , R.PRSNVRUS               '+#13#10+   //VERIFY ID
                '      , R.PRSNIPDT               '+#13#10+   //입력일자
                '      , R.PRSNIPTM               '+#13#10+   //입력시간
                '      , R.PRSNUSID               '+#13#10+   //입력자
                '      , P.ABHJNAME   PNM         '+#13#10+
                ' FROM MPSDTA.PRSNUMBM R          '+#13#10+
                '    , MASDTA.ABHJMSTM P          '+#13#10+
                ' where R.PRSNLBNO = '''+HBCD+''' '+#13#10+
                '   And R.PRSNMRNO = P.ABHJMRNO   ';
      try
          if SvrTEST then
              SQL.SaveToFile('오더.sql');

          Open;
      except
          on e:exception do begin
              TGlobal.ErrMsg:= e.Message+#13#10+SQL.Text;
              ShowMessage(e.Message);
              exit;
          end;
      end;

      if RecordCount=0 then exit;

      while Not Eof do begin
          ECD:= Trim(FieldByName('ECD').AsString);
          if TCode.IsSetCodeOK(ECD, TMaster.FIfCode) then
          begin
              TMaster.FPID:= Trim(FieldByName('PID').AsString);
              TMaster.FPNM:= FieldByName('PNM').AsString;
              TMaster.FDept:= FieldByName('DPT').AsString;
              TMaster.FIpDate:= FieldByName('IDT').AsString;
              TMaster.FOrdDate:= FieldByName('ODT').AsString;
              TMaster.PRSNORNO:= FieldByName('ONO').AsInteger;
              TMaster.PRSNORSQ:= FieldByName('OSeq').AsInteger;
              TMaster.FSLP := FieldByName('SLP').AsString;
              TMaster.FSUB := FieldByName('SUB').AsString;
              TMaster.FLAB := FieldByName('LAB').AsString;
              TMaster.FExamCode:= ECD;
              TMaster.FOrdState:= 'Y';
              Result:= True;
              exit;
          end;

          Next;
      end;
  end;
end;

function TDM.DownLoadOrder_QC_JEIL(TMaster: TIfMaster): boolean;
var
  BCD, ECD:string;
begin
  Result:= False;

  if SvrConnection = false then begin
      ShowMessage('로컬테스트중 입니다!');
      exit;
  end;

  BCD:= MakeBarCode(TMaster.FBarCode);

  with qrySOrder do begin
      Close;
      SQL.Text:=' SELECT QCSNRSLT AS Result   '+#13#10+
                '       ,QCSNMRNO AS PatNo    '+#13#10+
                '       ,QCSNLABR AS DEPT     '+#13#10+
                '       ,QCSNSEQN AS OrdSeq   '+#13#10+
                '       ,QCSNVSDT AS OrdDt    '+#13#10+
                '       ,QCSNCODE AS ExamCode '+#13#10+
                ' FROM MPSDTA.QCSNUMBM        '+#13#10+
                ' where QCSNLBNO = '''+BCD+''' ';

      try
          if SvrTEST then
              SQL.SaveToFile('QC오더.sql');

          Open;
      except
          on e:exception do begin
              TGlobal.ErrMsg:= e.Message+#13#10+SQL.Text;
              ShowMessage(e.Message);
              exit;
          end;
      end;

      if RecordCount=0 then exit;

      while Not Eof do begin
          ECD:= Trim(FieldByName('ExamCode').AsString);
          if TCode.IsSetCodeOK(ECD, TMaster.FIfCode) then
          begin
              TMaster.FPID:= FieldByName('PatNo').AsString;
              TMaster.FPNM:= FieldByName('QC').AsString;
              TMaster.PRSNORSQ:= FieldByName('OrdSeq').AsInteger;
              TMaster.FExamCode:= ECD;
              TMaster.FOrdState:= 'Y';
              Result:= True;
              exit;
          end;

          Next;
      end;
  end;
end;

function TDM.MakeBarCode(BCD: string): string;
begin
  Result:= BCD;

  if Copy(BCD,1,1) = 'C' then begin
      if Length(BCD) < 12 then
          Result:= 'C20'+Copy(BCD,2,Length(BCD)-1);
  end
  else begin
      if Length(BCD) < 12 then
          Result:= '20'+BCD;
  end;

end;

function TDM.UploadResult_Direct(TMaster: TIfMaster): boolean;
var
  UpCnt:integer;
  ECD, RES, FLG, UCD:string;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:= 'Local 테스트중입니다!';
      exit;
  end;

  if TMaster.FOrdState <> 'Y' then exit;

  if TMaster.FFlag = 'X' then exit;

  Result:= UploadHosp_One_SCHUH(TMaster);

end;

function TDM.UploadHospOne_JEIL_RES(TMaster:TIfMaster): boolean;
var
  HBCD, sNow, sDt, sTm:string;
  InstNo:integer;
begin
  Result:= False;

  if SvrConnection = False then begin
      ShowMessage('Local 테스트중입니다!');
      exit;
  end;

  sNow:= FormatDateTime('yyyymmddhhnnss', now);
  sDt:= Copy(sNow, 1, 8);
  sTm:= Copy(sNow, 9, 6);

  HBCD:= MakeBarCode(TMaster.FBarCode);

  with qrySUp do begin
      Close;
      SQL.Text:=' UPDATE MPSDTA.PRSNUMBM SET                  '+#13#10+
                '     PRSNRSLT   = '''+TMaster.FResult+'''    '+#13#10+  //결과값
                '    ,PRSNIPDT   = '''+sDt+'''                '+#13#10+  //결과입력일자
                '    ,PRSNIPTM   = '''+sTm+'''                '+#13#10+  //결과입력시간
                '    ,PRSNCASE   = :PRSNCASE                  '+#13#10+
                ' WHERE (PRSNRSLT = '''' or PRSNRSLT is Null) '+#13#10+  //결과값이 Null인 경우에만
                '   AND PRSNORNO  = :PRSNORNO                 '+#13#10+
                '   AND PRSNORSQ  = :PRSNORSQ                 '+#13#10+
                '   AND RTRIM(PRSNCODE)  = '''+TMaster.FExamCode+''' '+#13#10+
                '   AND PRSNMRNO  = '''+TMaster.FPID+'''      '+#13#10+   //Key  고객번호
                '   AND PRSNVSDP  = '''+TMaster.FDept+'''     '+#13#10+   //Key  내원과/재원과
                '   AND PRSNVSDT  = '''+TMaster.FIpDate+'''   '+#13#10+   //Key  내원일/입원일
                '   AND PRSNLDTE  = '''+TMaster.FOrdDate+'''  '+#13#10+   //Key  처방일자
                '   AND PRSNSLIP  = '''+TMaster.FSLP+'''      '+#13#10+   //Key  Slip코드
                //'   AND PRSNSUBC  = '''+TMaster.FSUB+'''      '+#13#10+   //Key  검사상세코드
                '   AND PRSNLBNO  = '''+HBCD+''' ';
      Parameters.ParamByName('PRSNORNO').Value:= TMaster.PRSNORNO;
      Parameters.ParamByName('PRSNORSQ').Value:= TMaster.PRSNORSQ;
      Parameters.ParamByName('PRSNCASE').Value:= GetInstNo(UpperCase(TGlobal.FIName), TGlobal.FICode);

      //ShowMessage(IntToStr(TMaster.PRSNORNO)+'|'+IntToStr(TMaster.PRSNORSQ)+'|'+IntToStr(GetInstNo(UpperCase(TGlobal.FIName), TGlobal.FICode)));

      try
          if SvrTEST then
              sql.SaveToFile('결과.sql');

          ExecSQL;

          Result:= UploadHospOne_JEIL_RMK(TMaster);
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= e.Message+#13#10+SQL.Text;
              ShowMessage(e.Message);
          end;
      end;
  end;

end;

function TDM.UploadHospOne_QC_JEIL(ADT, BCD, ECD, RES, RMK: string; OSEQ:integer): boolean;
var
  HBCD, sNow, sDt, sTm:string;
begin
  Result:= False;

  if SvrConnection = False then begin
      ShowMessage('Local 테스트중입니다!');
      exit;
  end;

  sNow:= FormatDateTime('yyyymmddhhnnss', now);
  sDt:= Copy(sNow, 1, 8);
  sTm:= Copy(sNow, 9, 6);

  HBCD:= MakeBarCode(BCD);

  with qrySUp do begin
      Close;
      SQL.Text:=' Update MPSDTA.QCSNUMBM Set    '+#13#10+
                '        QCSNRSLT = '''+RES+''' '+#13#10+
                //'      , QCSNCPMV = '''+RMK+''' '+#13#10+  //결과H/L(High/Low판정값)
                '      , QCSNRTDT = '''+sDt+''' '+#13#10+  //결과일자
                '      , QCSNRTTM = '''+sTm+''' '+#13#10+  //결과시간
                ' Where QCSNLBNO = '''+HBCD+''' '+#13#10+
                '   And RTrim(QCSNCODE) = '''+ECD+'''  '+#13#10+
                '   And QCSNVSDT = '''+ADT+'''  '+#13#10+
                '   And QCSNSEQN = :QCSNSEQN    ';
      Parameters.ParamByName('QCSNSEQN').Value:= OSEQ;
      try
          if SvrTEST then
              sql.SaveToFile('QC결과.sql');

          ExecSQL;
          Result:= True;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= e.Message+#13#10+SQL.Text;
              ShowMessage(e.Message);
          end;
      end;
  end;

end;

function TDM.GetInstNo(INM, ICD: string): integer;
var
  INO:integer;
begin
  Result:= 0;

  if INM = 'CHORUSTRIO' then
      INO:= StrToIntDef(ICD, 341)
  else
  if INM = 'TEST1' then
      INO:= StrToIntDef(ICD, 131)
  else
      INO:= StrToIntDef(ICD, 0);

  Result:= INO;
end;

function TDM.UploadHospOne_JEIL_RMK(TMaster: TIfMaster): boolean;
var
  sNow, sDt, sTm:string;
begin
  Result:= False;

  if SvrConnection = False then begin
      ShowMessage('Local 테스트중입니다!');
      exit;
  end;

  //리마크가 없으면 안넣는다..
  if TMaster.FRMK = '' then begin
      Result:= True;
      exit;
  end;

  sNow:= FormatDateTime('yyyymmddhhnnss', now);
  sDt:= Copy(sNow, 1, 8);
  sTm:= Copy(sNow, 9, 6);

  with qrySUp do begin
      Close;
      SQL.Text:=' Select Count(PRSLMRNO) Cnt                 '+#13#10+
                ' From MPSDTA.PRSLRMKM                       '+#13#10+
                ' Where PRSLMRNO = '''+TMaster.FPID+'''      '+#13#10+
                '   And PRSLVSDP = '''+TMaster.FDept+'''     '+#13#10+
                '   And PRSLVSDT = '''+TMaster.FIpDate+'''   '+#13#10+
                '   And PRSLORNO = :PRSLORNO                 '+#13#10+
                '   And PRSLORSQ = :PRSLORSQ                 ';
      Parameters.ParamByName('PRSLORNO').Value:= TMaster.PRSNORNO;
      Parameters.ParamByName('PRSLORSQ').Value:= TMaster.PRSNORSQ;
      try
          if SvrTEST then
              sql.SaveToFile('리마크조회.sql');

          Open;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= e.Message+#13#10+SQL.Text;
              ShowMessage(e.Message);
          end;
      end;

      if Fields[0].AsInteger > 0 then begin
          SQL.Text:=' Update MPSDTA.PRSLRMKM                    '+#13#10+
                    '  set PRSLIPDT = '''+sDt+'''               '+#13#10+ //===>Comment입력(수정)일자
                    '     ,PRSLIPTM = '''+sTm+'''               '+#13#10+ //===>Comment입력(수정)시간
                    //'     ,PRSLUSID = 'ED002'                   '+#13#10+ //===>Comment입력자(수정자)ID
                    '     ,PRSLRMRK = '''+TMaster.FRMK+'''      '+#13#10+ //->Comment내용
                    '  Where PRSLMRNO = '''+TMaster.FPID+'''    '+#13#10+ //  ===>고객번호
                    '    And PRSLVSDP = '''+TMaster.FDept+'''   '+#13#10+ //        ===>진료과
                    '    And PRSLVSDT = '''+TMaster.FIpDate+''' '+#13#10+ //  ===>내원일자
                    '    And PRSLORNO = :PRSLORNO               '+#13#10+        //+#13#10+ // ===>Order No
                    '    And PRSLORSQ = :PRSLORSQ               '; //===>Order Seq
          Parameters.ParamByName('PRSLORNO').Value:= TMaster.PRSNORNO;
          Parameters.ParamByName('PRSLORSQ').Value:= TMaster.PRSNORSQ;
          try
              if SvrTEST then
                  sql.SaveToFile('리마크UPDATE.sql');

              ExecSQL;
              Result:= True;
          except
              on e:exception do
              begin
                  TGlobal.ErrMsg:= e.Message+#13#10+SQL.Text;
                  ShowMessage(e.Message);
              end;
          end;
      end
      else begin
          //PRSLUSID 입력자
          SQL.Text:=' Insert Into MPSDTA.PRSLRMKM                         '+#13#10+
                    '             (PRSLMRNO,PRSLVSDP,PRSLVSDT,PRSLORNO    '+#13#10+
                    '             ,PRSLCASE,PRSLIPDT,PRSLIPTM,PRSLRMRK, PRSLORSQ)   '+#13#10+
                    ' Values ('''+TMaster.FPID+'''                        '+#13#10+
                    '        ,'''+TMaster.FDept+'''                       '+#13#10+
                    '        ,'''+TMaster.FIpDate+'''                     '+#13#10+
                    '        ,:PRSLORNO                                   '+#13#10+
                    '        ,0                                           '+#13#10+
                    '        ,'''+sDt+'''                                 '+#13#10+
                    '        ,'''+sTm+'''                                 '+#13#10+
                    '        ,'''+TMaster.FRMK+'''                        '+#13#10+
                    '        ,:PRSLORSQ )                                 ';

          Parameters.ParamByName('PRSLORNO').Value:= TMaster.PRSNORNO;
          Parameters.ParamByName('PRSLORSQ').Value:= TMaster.PRSNORSQ;
          try
              if SvrTEST then
                  sql.SaveToFile('리마크INSERT.sql');

              ExecSQL;
              Result:= True;
          except
              on e:exception do
              begin
                  TGlobal.ErrMsg:= e.Message+#13#10+SQL.Text;
                  ShowMessage(e.Message);
              end;
          end;
      end;
  end;
end;

function TDM.DownLoadOrder_JND_RES(TMaster: TIfMaster): boolean;
var
  i, nCnt:integer;
  ECD, OutMsg, Pas:string;
begin
  Result:= False;
  exit;

  {
  TMaster.FOrdState:= 'Y';
  Result:= True;
  TMaster.FExamCode:= 'LIS94';
  exit;
  }
  {
  if DownLoadTux_Str(TMaster.FBarCode, OutMsg) = true then begin
      nCnt:= CountStr(OutMsg, ETX);
      for i:=1 to nCnt do begin
          Pas:= TokenStr(OutMsg, ETX, i);
          TMaster.FPID:= TokenStr(Pas,'|',1);
          TMaster.FPNM:= TokenStr(Pas,'|',2);
          ECD:= TokenStr(Pas,'|',4);

          if TCode.GetExamCode(TMaster.FIfCode) = ECD then begin
              TMaster.FExamCode:= ECD;
              TMaster.FOrdState:= 'Y';
              Result:= True;
              exit;
          end;
      end;
  end
  else begin
      TMaster.FExamCode:= TCode.GetExamCode(ECd);
      TMaster.FOrdState:= 'N';
  end;  }
end;

function TDM.UploadHosp_One_JND(TMaster:TIfMaster): boolean;
var
  SvrMsg:string;
begin
  Result:= False;
{  if Not UpLoadTux_Str(TMaster.FBarCode, TMaster.FExamCode, TMaster.FResult, SvrMsg) then
      ShowMessage(SvrMsg)
  else
      Result:= True;
}
{
  Result:= False;
  OCP034LA 서비스를 태우시고
input값은
acptacdt 바코드 접수일자
acptsrno 검체번호
rslnstat 상태값 = 'T'
rslnuser 사용자 = 현재 미정(상의후 결정)
acptitem 대표 검체코드 'LIS'
ocp_selcnt 검체개수
pseudo10 장비코드 = 'h'
rslnitem 검체개수만큼의 검체코드
}
end;

procedure TDM.DataModuleDestroy(Sender: TObject);
begin
  //TuxTerm;
end;

function TDM.CheckLot(INm, Enm, Lot, Typ, Lev: string): boolean;
begin
  Result:= False;

  with qryC1 do begin
      Close;
      SQL.Text:= ' Select * From TB_Lot '+
                 ' Where ENM = '''+ENM+''' '+
                 '   And Lot = '''+LOT+''' '+
                 '   And IName = '''+INm+''' '+
                 '   and Typ = '''+Typ+''' '+
                 '   and Lev = '''+Lev+''' ';
      Open;

      if RecordCount > 0 then
          Result:= True;
  end;

end;

procedure TDM.SaveOneLotInfo(INm, LotNm, Lev, ENM, Mean, SD, Fdt, Tdt,
  sLow, sHigh, Typ: string);
begin
  with dm.qryC2 do begin
      Close;
      SQL.Text:= ' Select * From tb_Lot ' +
                 ' Where INAME = '''+iNM+''' '+
                 '   and Lot = '''+LotNm+''' '+
                 '   and Lev = '''+Lev+''' '+
                 '   and ENM = '''+ENM+''' '+
                 '   and Typ = '''+Typ+''' ';
      Open;

      if RecordCount = 0 then begin
          Close;
          SQL.Text:= ' Insert Into Tb_Lot (IName, Lot, Lev, Enm, mean, sd, fdt, tdt, r_Low, r_High, Typ) '+
                     ' Values ( '''+iNM+''', '''+LotNm+''', '''+Lev+''', '''+ENM+''', '+
                     ' '+Mean+', '+SD+', '''+fdt+''', '''+Tdt+''','+sLow+','+sHigh+','''+Typ+''') ';
          ExecSql;
      end
      else begin
          Close;
          SQL.Text:= ' Update Tb_Lot Set            '+
                     '     mean = '+Mean+'      '+
                     '   , sd   = '+SD+'        '+
                     '   , fdt  = '''+fdt+'''       '+
                     '   , tdt  = '''+Tdt+'''       '+
                     '   , r_Low= '+sLow+'      '+
                     '   , r_High= '+sHigh+'     '+
                     ' Where INAME = '''+iNM+''' '+
                     '   and Lot = '''+LotNm+''' '+
                     '   and Lev = '''+Lev+''' '+
                     '   and ENM = '''+ENM+''' '+
                     '   and Typ = '''+Typ+''' ';
          ExecSql;
      end;
  end;

end;

procedure TDM.SaveQC(TMaster: TIfMaster);
var
  ESeq:string;
begin
  ESeq:= GetExamSeq_QC(TMaster.FExamDate);

  with DM.qryC2 do begin
     Close;
     SQL.Text:= ' Select * From tb_QC '+
                ' Where ExamDate = '''+TMaster.FExamDate+''' '+
                '   And ExamSeq = '''+ESeq+'''  ';
     Open;

     if RecordCount > 0 then begin
         Close;
         SQL.Text:= ' Update tb_QC Set '+
                    '      result = '''+TMaster.FResult+'''     '+
                    '    , TYP    = '''+TMaster.FTYP+'''     '+
                    '    , Lot    = '''+TMaster.FLotNo+'''     '+
                    '    , ExamTime = '''+TMaster.FRcvTime+'''   '+
                    '    , UpCode   = '''+TMaster.FIfCode+'''   '+
                    '    , Comment = '''+TMaster.FReMark+'''    '+
                    '    , barcode = '''+TMaster.FBarCode+'''   '+
                    '    , Lev     = '''+TMaster.FLotLev+'''    '+
                    ' where ExamDate =  '''+TMaster.FExamDate+''' '+
                    '   And ExamSeq = '''+ESeq+'''   ';
     end
     else begin
         Close;
         SQL.Text:= ' Insert Into tb_QC (ExamDate, ExamSeq, TYP, ExamTime, Result, Lot, UpCode, Comment, Lev, barcode) '+
                    ' Values '+
                    ' ('''+TMaster.FExamDate+''', '''+ESeq+''', '''+TMaster.FTYP+''', '''+TMaster.FRcvTime+''','''+TMaster.FResult+''', '+
                    '  '''+TMaster.FLotNo+''', '''+TMaster.FIfCode+''', '''+TMaster.FReMark+''','''+TMaster.FLotLev+''', '''+TMaster.FBarCode+''' )';
     end;

     ExecSql;
  end;

end;

procedure TDM.SetExamLotMean(TMaster: TIfMaster);
var
  DigIdx, DigCnt:integer;
  m, Sd:double;
  dMin, dMax:double;
begin
  dMin:=0; dMax:=0;

  DigIdx:= Pos('.', TMaster.FResult);
  DigCnt:= Length( Copy(TMaster.FResult, DigIdx+1, Length(TMaster.FResult)-DigIdx) );

  Case DigCnt of
      0:begin
            dMin:= StrToFloat(TMaster.FLotMin);
            dMax:= StrToFloat(TMaster.FLotMax);
        end;
      1:begin
            dMin:= StrToFloat(TMaster.FLotMin) / 10;
            dMax:= StrToFloat(TMaster.FLotMax) / 10;
        end;
      2:begin
            dMin:= StrToFloat(TMaster.FLotMin) / 100;
            dMax:= StrToFloat(TMaster.FLotMax) / 100;
        end;
      3:begin
            dMin:= StrToFloat(TMaster.FLotMin) / 1000;
            dMax:= StrToFloat(TMaster.FLotMax) / 1000;
        end;
  end;

  TMaster.FLotMin:= Trim(Format('%5.'+IntToStr(DigCnt)+'f', [dMin]));
  TMaster.FLotMax:= Trim(Format('%5.'+IntToStr(DigCnt)+'f', [dMax]));

  m:= mean([dMin, dMax]);
  sd:= stddev([dMin, dMax]);


  TMaster.FLotMean:= Trim(Format('%5.'+IntToStr(DigCnt)+'f', [m]));
  TMaster.FLotSD:= Trim(Format('%5.'+IntToStr(DigCnt)+'f', [sd]));

end;

function TDM.GetExamSeq_QC(ExamDate: string): string;
var
  TSql: TQueryInfo;
  QryEx: TAdoQuery;
  Seq:integer;
begin
  Result:= '001';
  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Application);

  try
      with TSql do begin
        Clear;
        AddSql(' Select iif( isNull(Max(Val(ExamSeq))), 1, Max(Val(ExamSeq))+1) As SEQ From TB_Qc ');
        AddSql(' Where ExamDate = '''+ExamDate+'''         ');
        RCount:= LocalSelect(QryEx);

        if RCount > 0 then begin
            Seq:= QryEx.Fields[0].AsInteger;
            Result:= PadLeftStr(IntToStr(Seq), '0', 3);
        end;
    end;

  finally
      TSql.Free;
      QryEx.Free;
  end;

end;

function TDM.DownLoadOrder_CBD(TMaster: TIfMaster): boolean;
var
  ECD:string;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:=  'Local TEST중입니다.';
      exit;
  end;

  with qrySOrder do begin
      Close;
      SQL.Text:=' select                                         '+#13#10+
                '   M.ID                                           '+#13#10+
                ' , M.PTNO                                         '+#13#10+
                ' , M.ORDERID                                      '+#13#10+
                ' , M.SNAME                                        '+#13#10+
                ' , M.SEX                                          '+#13#10+
                ' , M.AGE                                          '+#13#10+
                ' , M.IPDOPD                                       '+#13#10+
                ' , M.DEPTCODE                                     '+#13#10+
                ' , M.SPECIMENNUM                                  '+#13#10+
                ' , M.ORDERCODE                                    '+#13#10+
                ' , E.CODENAME                                     '+#13#10+
                ' , E.SUBCODE                                      '+#13#10+
                ' , E.SUBCODESEQ                                   '+#13#10+
                ' , E.UNIT                                         '+#13#10+
                ' , E.REFV                                         '+#13#10+
                ' from  LABRECEPTION M, LABCODE E                  '+#13#10+
                ' where M.SPECIMENNUM = '''+TMaster.FBarCode+'''   '+#13#10+
                '   And M.JSTATUS <= ''3''                         '+#13#10+
                '   And M.ORDERCODE = E.ORDERCODE                  '+#13#10+
                '   And M.DESTINATION1 = E.DESTINATION1            '+#13#10+
                '   And M.DESTINATION2 = E.DESTINATION2            '+#13#10;
      try
          if SvrTEST = True then
              SQL.SaveToFile('.\오더조회.sql');
              
          Open;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= 'LABRECEPTION 오더체크 에러 입니다! 에러메세지->'+ e.Message;
              ShowMessage(e.Message);
              exit;
          end;
      end;

      if RecordCount > 0 then begin
          while Not Eof do begin
              ECD:= FieldByName('SUBCODE').AsString;
              if TCode.IsSetCodeOK(Trim(ECD), TMaster.FIfCode) then begin
                  TMaster.FOrdCode  := FieldByName('ORDERCODE').AsString;;
                  TMaster.FOrdName  := FieldByName('CODENAME').AsString;
                  TMaster.FExamCode := ECD;
                  TMaster.SUBCODESEQ:= FieldByName('SUBCODESEQ').AsString;
                  TMaster.CodeUnit  := FieldByName('UNIT').AsString;
                  TMaster.REFV      := FieldByName('REFV').AsString;
                  TMaster.FPID  := FieldByName('PTNO').AsString;
                  TMaster.FPNM  := FieldByName('SName').AsString;
                  //TMaster.FAge  := FieldByName('AGE').AsString;
                  //TMaster.FSex  := FieldByName('SEX').AsString;
                  TMaster.FLAB  := IntToStr(FieldByName('ID').AsInteger);
                  Result:= True;
                  TMaster.FOrdState:= 'Y';
                  Exit;
              end;
              Next;
          end;
      end;
  end;
end;

function TDM.UploadHosp_One_CBD(TMaster: TIfMaster): boolean;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:= 'Local 테스트중입니다!';
      exit;
  end;

  if TMaster.FOrdState <> 'Y' then exit;

  with qrySUp do begin
      Close;
      SQL.Text:=' Select * From LabResult                        '+#13#10+
                ' where SPECIMENNUM = '''+TMaster.FBarCode+'''    '+#13#10+
                '   And ID = '+TMaster.FLAB+'                    '+#13#10+
                '   And ORDERCODE = '''+TMaster.FOrdCode+'''     '+#13#10+
                '   And SUBCODE   = '''+TMaster.FExamCode+'''    ';

      if SvrTEST = True then
          SQL.SaveToFile('.\결과전송조회.sql');

      Open;

      if RecordCount = 0 then begin
          Close;
          SQL.Text:=' Insert Into LabResult (ID, ORDERCODE, SUBCODE, SUBCODENAME, SEQ, UNITCODE, RDATA, SPECIMENNUM, RESULTDATE, RESULTTIME) '+#13#10+
                    ' Values '+#13#10+
                    ' ('+TMaster.FLAB+','''+TMaster.FOrdCode+''','''+TMaster.FExamCode+''','''+TMaster.FOrdName+''', '+TMaster.SUBCODESEQ+', '+
                    '  '''+TMaster.CodeUnit+''','''+TMaster.FResult+''','''+TMaster.FBarCode+'''  '+
                    '  , TRUNC(SYSDATE), To_Char(SysDate, ''HH24:MI:SS'') ) ';
      end
      else begin
          Close;
          SQL.Text:=' update LABRESULT set RDATA = '''+TMaster.FResult+'''       '+
                    '                    , ResultDate = TRUNC(SYSDATE) '+#13#10+
                    '                    , ResultTime = To_Char(SysDate, ''HH24:MI:SS'')  '+#13#10+
                    ' where SPECIMENNUM = '''+TMaster.FBarCode+'''    '+#13#10+
                    '   And ID = '+TMaster.FLAB+'                    '+#13#10+
                    '   And ORDERCODE = '''+TMaster.FOrdCode+'''     '+#13#10+
                    '   And SUBCODE   = '''+TMaster.FExamCode+'''     ';
      end;

      if SvrTEST = True then
          SQL.SaveToFile('.\결과전송.sql');

      try
          ExecSQL;
          Result:= True;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= 'LABRESULT 결과전송 에러 입니다! 에러메세지->'+ e.Message;
              ShowMessage(e.Message);
          end;
      end;
  end;

end;

function TDM.UploadHosp_STATE_CBD(BCD, OCD: string): boolean;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:= 'Local 테스트중입니다!';
      exit;
  end;

  with qrySUp do begin
      Close;
      SQL.Text:=' UPDATE LABRECEPTION SET JSTATUS=''3'', RESULTDATE=TRUNC(SYSDATE) '+
                ' Where SPECIMENNUM =  '''+BCD+''' '+
                '   And ORDERCODE   = '''+OCD+''' '+
                '   And JSTATUS <= ''3'' ';
                
      if SvrTEST = True then
          SQL.SaveToFile('.\상태변경.sql');

      try
          ExecSQL;
          Result:= True;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= 'LABRECEPTION 상태변경 에러 입니다! 에러메세지->'+ e.Message;
              ShowMessage(e.Message);
          end;
      end;
  end;

end;

function TDM.DownLoadOrder_SCHUH(TMaster: TIfMaster): boolean;
var
  K:integer;
  ECD:string;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:=  'Local TEST중입니다.';
      exit;
  end;

  with qrySOrder do begin
      Close;
      SQL.Text:=' select M.PART_JUBSU_DATE                       '+#13#10+
                '      , M.PART_JUBSU_TIME                       '+#13#10+
                '      , M.BUNHO                                 '+#13#10+
                '      , M.SUNAME                                '+#13#10+
                '      , M.AGE                                   '+#13#10+
                '      , M.SEX                                   '+#13#10+
                '      , M.SPECIMEN_CODE                         '+#13#10+
                '      , M.GWA_NAME                              '+#13#10+
                '      , M.HANGMOG_CODE                          '+#13#10+
                '      , E.GUMSA_NAME                            '+#13#10+
                '      , E.JANGBI_OUT_CODE                       '+#13#10+
                '      , E.JANGBI_CODE                           '+#13#10+
                '      , R.CONFIRM_YN                            '+#13#10+
                '      , R.CPL_RESULT                            '+#13#10+
                '      , R.JANGBI_YN                             '+#13#10+
                '      , R.JANGBI_CODE                           '+#13#10+
                ' from CPL3020 R                                 '+#13#10+
                '    , CPL2010 M                                 '+#13#10+
                '    , CPL0101 E                                 '+#13#10+
                ' where R.SPECIMEN_SER ='''+TMaster.FBarCode+''' '+#13#10+
                '   and NVL(R.CONFIRM_YN, ''N'') = ''N''         '+#13#10+
                '   and R.JANGBI_CODE = '''+TGlobal.FICode+'''   '+#13#10+
                '   and E.JANGBI_OUT_CODE is Not Null            '+#13#10+
                '   and R.SPECIMEN_SER = M.SPECIMEN_SER          '+#13#10+
                '   and R.HANGMOG_CODE = M.HANGMOG_CODE          '+#13#10+
                '   and R.SPECIMEN_CODE = M.SPECIMEN_CODE        '+#13#10+
                '   and R.HANGMOG_CODE = E.HANGMOG_CODE          '+#13#10+
                '   and R.SPECIMEN_CODE = E.SPECIMEN_CODE        ';
                {
                //테스트용..
                '   and R.JANGBI_CODE = '''+TGlobal.FICode+'''   '+#13#10+
                '   and E.JANGBI_OUT_CODE is Not Null            '+#13#10+
                '   and R.HANGMOG_CODE = ''LS2795''              '+#13#10+}



      try
          if SvrTEST = True then
              SQL.SaveToFile('.\오더조회.sql');

          Open;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= ' 오더체크 에러 입니다! 에러메세지->'+ e.Message;
              ShowMessage(e.Message);
              exit;
          end;
      end;

      if RecordCount > 0 then begin
          TMaster.vOrder:= VarArrayCreate([0, RecordCount-1], varVariant);
          TMaster.vAbbr := VarArrayCreate([0, RecordCount-1], varVariant);
          TMaster.vOrdList:= VarArrayCreate([0, RecordCount-1], varVariant);
          TMaster.vUpCode:=  VarArrayCreate([0, RecordCount-1], varVariant);
          K:=0;

          while Not Eof do begin
              ECD:= FieldByName('JANGBI_OUT_CODE').AsString;
              //테스트용..
              //ECD:= FieldByName('HANGMOG_CODE').AsString;
              if TCode.SetCode_ECode(Trim(ECD)) then begin
                  TMaster.vOrder[K]:= TCode.GetIfCode(ECD);
                  TMaster.vAbbr[K] := TCode.GetAbbr(ECD);
                  TMaster.vOrdList[K]:= ECD;
                  TMaster.vUpCode[K]:= TCode.GetUpCode(ECD);
                  Inc(K);

                  if Result = False then begin
                      TMaster.FOrdCode  := FieldByName('HANGMOG_CODE').AsString;;
                      TMaster.FPID  := FieldByName('BUNHO').AsString;
                      TMaster.FPNM  := FieldByName('SUNAME').AsString;
                      Result:= True;
                      TMaster.FOrdState:= 'Y';
                  end;
              end;

              Next;
          end;
      end;
  end;

{
                ' SELECT B.BUNHO                                            '+#13#10+
                '      , B.SUNAME                                           '+#13#10+
                '      , B.AGE                                              '+#13#10+
                '      , B.SEX                                              '+#13#10+
                '      , B.GWA_NAME AS DEPT                                 '+#13#10+
                '      , B.HANGMOG_CODE                                     '+#13#10+
                '      , B.SPECIMEN_CODE                                    '+#13#10+
                '      , B.PART_JUBSU_DATE                                  '+#13#10+
                '      , B.PART_JUBSU_TIME                                  '+#13#10+
                '      , B.Doctor_Name                                      '+#13#10+
                '      , A.Lab_No                                           '+#13#10+
                '      , B.GUMSA_NAME                                       '+#13#10+
                '      , B.JANGBI_OUT_CODE                                  '+#13#10+
                '       FROM CPL2010 B,                                     '+#13#10+
                '            CPL3010 A,                                     '+#13#10+
                '            CPL0101 C                                      '+#13#10+
                '      WHERE B.PART_JUBSU_DATE = A.PART_JUBSU_DATE          '+#13#10+
                '        AND B.JUNDAL_GUBUN    = A.JUNDAL_GUBUN             '+#13#10+
                '        AND B.SPECIMEN_SER    = A.SPECIMEN_SER             '+#13#10+
                '        AND B.JANGBI_CODE     = '''+TGlobal.FICode+'''     '+#13#10+
                '        AND B.SPECIMEN_SER    = '''+TMaster.FBarCode+'''   '+#13#10;
}
end;

function TDM.UploadHosp_One_SCHUH(TMaster: TIfMaster): boolean;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:= 'Local 테스트중입니다!';
      exit;
  end;

  if TMaster.FOrdState <> 'Y' then exit;

  with spUp do begin
      Close;
      Parameters.ParamByName('I_USER_ID').Value:= TGlobal.FUserID;
      Parameters.ParamByName('I_JANGBI_CODE').Value:= TGlobal.FICode;
      Parameters.ParamByName('I_SPECIMEN_SER').Value:= TMaster.FBarCode;
      Parameters.ParamByName('I_JANGBI_OUT_CODE').Value:= TMaster.FUpCode;
      Parameters.ParamByName('I_CPL_RESULT').Value:= TMaster.FResult;
      Parameters.ParamByName('I_RESULT_DATE').Value:= FormatDateTime('yyyymmdd', now);
      Parameters.ParamByName('I_RESULT_SEQ').Value:= '';
      try
          ExecProc;
          Result:= True;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= '결과전송 에러 입니다! 에러메세지->'+ e.Message;
              ShowMessage(e.Message);
          end;
      end;
  end;
end;

function TDM.DownLoadOrder_SCHUH_One(TMaster: TIfMaster): boolean;
var
  ECD:string;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:=  'Local TEST중입니다.';
      exit;
  end;

  with qrySOrder do begin
      Close;
      SQL.Text:=' select M.PART_JUBSU_DATE                       '+#13#10+
                '      , M.PART_JUBSU_TIME                       '+#13#10+
                '      , M.BUNHO                                 '+#13#10+
                '      , M.SUNAME                                '+#13#10+
                '      , M.AGE                                   '+#13#10+
                '      , M.SEX                                   '+#13#10+
                '      , M.SPECIMEN_CODE                         '+#13#10+
                '      , M.GWA_NAME                              '+#13#10+
                '      , M.HANGMOG_CODE                          '+#13#10+
                '      , E.GUMSA_NAME                            '+#13#10+
                '      , E.JANGBI_OUT_CODE                       '+#13#10+
                '      , E.JANGBI_CODE                           '+#13#10+
                '      , R.CONFIRM_YN                            '+#13#10+
                '      , R.CPL_RESULT                            '+#13#10+
                '      , R.JANGBI_YN                             '+#13#10+
                '      , R.JANGBI_CODE                           '+#13#10+
                ' from CPL3020 R                                 '+#13#10+
                '    , CPL2010 M                                 '+#13#10+
                '    , CPL0101 E                                 '+#13#10+
                ' where R.SPECIMEN_SER ='''+TMaster.FBarCode+''' '+#13#10+
                '   and NVL(R.CONFIRM_YN, ''N'') = ''N''         '+#13#10+
                '   and R.JANGBI_CODE = '''+TGlobal.FICode+'''   '+#13#10+
                '   and E.JANGBI_OUT_CODE is Not Null            '+#13#10+
                '   and R.SPECIMEN_SER = M.SPECIMEN_SER          '+#13#10+
                '   and R.HANGMOG_CODE = M.HANGMOG_CODE          '+#13#10+
                '   and R.SPECIMEN_CODE = M.SPECIMEN_CODE        '+#13#10+  
                '   and R.HANGMOG_CODE = E.HANGMOG_CODE          '+#13#10+
                '   and R.SPECIMEN_CODE = E.SPECIMEN_CODE        ';

      try
          if SvrTEST = True then
              SQL.SaveToFile('.\오더조회.sql');

          Open;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= ' 오더체크 에러 입니다! 에러메세지->'+ e.Message;
              ShowMessage(e.Message);
              exit;
          end;
      end;

      if RecordCount > 0 then begin
          while Not Eof do begin
              ECD:= FieldByName('JANGBI_OUT_CODE').AsString;
              if TCode.IsSetCodeOK(Trim(ECD), TMaster.FIfCode) then begin
                  TMaster.FOrdCode  := FieldByName('HANGMOG_CODE').AsString;;
                  TMaster.FExamCode := ECD;
                  TMaster.FPID  := FieldByName('BUNHO').AsString;
                  TMaster.FPNM  := FieldByName('SUNAME').AsString;
                  //TMaster.FLAB  := FieldByName('Lab_No').AsString;
                  Result:= True;
                  TMaster.FOrdState:= 'Y';
                  Exit;
              end;

              Next;
          end;
      end;
  end;
end;

procedure TDM.SetDownCode(TMaster: TIfMaster);
var
  i, VC:integer;
  UpCD, ECD, OrdSeq:string;
begin
  TMaster.IsDownCodeOK:= False;
  TMaster.FExamCode:= '';
  TMaster.FOrdSeq  := '';

  VC:= VarArrayDimCount(TMaster.vUpCode);
  if VC > 0 then begin
      for i:=0 to VarArrayHighBound(TMaster.vUpCode,1) do begin
         UpCD:= Trim(TMaster.vUpCode[i]);

         if UpCD = TMaster.FUpCode then begin
             TMaster.FOrdState:= 'Y';
             TMaster.IsDownCodeOK:= True;
             TMaster.FExamCode:= TMaster.vOrdList[i];
             TMaster.FOrdCode := TMaster.vOrdCdList[i];
             TMaster.FANO     := TMaster.vANO[i]; 
             exit;
         end;
      end;
  end;

  //다운받은코드가 없을때!
  TMaster.FExamCode:= TCode.GetExamCode_UpCode(TMaster.FUpCode);
end;

function TDM.DownLoadOrder_DJ(TMaster: TIfMaster): boolean;
var
  K:integer;
  ECD:string;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:=  'Local TEST중입니다.';
      exit;
  end;

  with qrySOrder do begin
      Close;
      SQL.Text:=' Select          M.IpdOpd GBN                   '+#13#10+
                '               , M.ReceiptDate ADT              '+#13#10+
                '               , M.ReceiptNo ANO                '+#13#10+
                '               , M.OrderCode OCD                '+#13#10+
                '               , R.LabCode  ECD                 '+#13#10+
                '               , M.PTno PID                     '+#13#10+
                '               , M.SName PNM                    '+#13#10+
                '               , M.Age                          '+#13#10+
                '               , M.ReceiptTime                  '+#13#10+
                '               , M.Sex                          '+#13#10+
                '               , M.BI                           '+#13#10+
                '               , M.DeptCode                     '+#13#10+
                '               , M.WardCode                     '+#13#10+
                '               , M.Roomcode                     '+#13#10+
                '               , M.BillFlag                     '+#13#10+
                '               , M.JStatus                      '+#13#10+
                '               , M.SPECIMENNUM                  '+#13#10+
                ' From SLA_LabMaster M                           '+#13#10+
                '    , SLA_LABRESULT R                           '+#13#10+
                ' Where M.ReceiptNo = R.ReceiptNo                '+#13#10+
                '   And M.ORDERCODE = R.ORDERCODE                '+#13#10+
                '   And R.SPECIMENNUM = '''+TMaster.FBarCode+''' '+#13#10+
                '   And M.JsTATUS < ''3''                        '+#13#10;


      try
          if SvrTEST = True then
              SQL.SaveToFile('.\오더조회.sql');

          Open;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= ' 오더체크 에러 입니다! 에러메세지->'+ e.Message;
              ShowMessage(e.Message);
              exit;
          end;
      end;

      if RecordCount > 0 then begin
          TMaster.vOrder:= VarArrayCreate([0, RecordCount-1], varVariant);
          TMaster.vAbbr := VarArrayCreate([0, RecordCount-1], varVariant);
          TMaster.vOrdList:= VarArrayCreate([0, RecordCount-1], varVariant);
          TMaster.vUpCode:=  VarArrayCreate([0, RecordCount-1], varVariant);
          TMaster.vOrdCdList:=  VarArrayCreate([0, RecordCount-1], varVariant);
          TMaster.vANO:=  VarArrayCreate([0, RecordCount-1], varVariant);
          K:=0;

          while Not Eof do begin
              ECD:= Trim(FieldByName('ECD').AsString);

              //ShowMessage(ECD);

              if TCode.SetCode_ECode(ECD) then begin
                  TMaster.vOrder[K]:= TCode.GetIfCode(ECD);
                  TMaster.vAbbr[K] := TCode.GetAbbr(ECD);
                  TMaster.vOrdList[K]:= FieldByName('ECD').AsString;
                  TMaster.vUpCode[K]:= TCode.GetUpCode(ECD);
                  TMaster.vOrdCdList[K]:= FieldByName('OCD').AsString;
                  TMaster.vANO[K]:= FieldByName('ANO').AsString;
                  Inc(K);

                  if Result = False then begin
                      TMaster.FOrdCode  := FieldByName('OCD').AsString;
                      TMaster.FPID  := FieldByName('PID').AsString;
                      TMaster.FPNM  := FieldByName('PNM').AsString;
                      TMaster.FADT  := FieldByName('ADT').AsString;
                      TMaster.FANO  := FieldByName('ANO').AsString;
                      Result:= True;
                      TMaster.FOrdState:= 'Y';
                  end;
              end;

              Next;
          end;
      end;
  end;
end;

function TDM.UploadHosp_One_DJ(TMaster:TIfMaster): boolean;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:= 'Local 테스트중입니다!';
      exit;
  end;

  with qrySUp do begin
      Close;
      SQL.Text:=' update SLA_LABRESULT set Result = '''+TMaster.FResult+'''              '+#13#10+
                '                        , TransFlag = ''1''                             '+#13#10+
                '                        , ResultDate = To_Char(SysDate, ''YYYY-MM-DD'') '+#13#10+
                '                        , ResultTime = To_Char(SysDate, ''HH24:MI:SS'') '+#13#10+
                ' where ReceiptNo = '+TMaster.FANO+'                                     '+#13#10+
                '   And ORDERCODE = '''+TMaster.FOrdCode+'''    '+#13#10+
                '   And LABCODE   = '''+TMaster.FExamCode+'''   '+#13#10+
                '   And transflag < ''2'' ';

      if SvrTEST = True then
          SQL.SaveToFile('.\결과등록.sql');

      try
          ExecSQL;
          Result:= True;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= 'SLA_LABRESULT 결과전송 에러 입니다! 에러메세지->'+ e.Message;
              ShowMessage(e.Message);
          end;
      end;
  end;

end;

function TDM.UploadState_DJ(TMaster:TIfMaster): boolean;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:=  'Local TEST중입니다.';
      exit;
  end;


  with qrySUp1 do begin
      Close;
      SQL.Text:= ' Update SLA_LABMASTER SET JSTATUS = ''3'' '+
                 ' Where ReceiptNo = '+TMaster.FANO+
                 '   And OrderCode = '''+TMaster.FOrdCode+''' '+
                 '   And JStatus < ''3'' ';
      try
          ExecSQL;
          Result:= True;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= 'SLA_LABMASTER JSTATUS변경 에러입니다! 에러메세지->'+ e.Message;
          end;
      end;
  end;

end;

function TDM.DownLoadOrder_DJI(TMaster: TIfMaster): boolean;
var
  K:integer;
  ECD:string;
  ODT:string;
  FSEQ, TSEQ:string;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:=  'Local TEST중입니다.';
      exit;
  end;

  //if Length(TMaster.FBarCode) < 11 then exit;

  ODT:= '20'+ Copy(TMaster.FBarCode, 1, 6);

  //인덱스 안타므로 SEQ를 미리 만들자..
  FSEQ:= ODT + '0000000';
  TSEQ:= ODT + '1000000';

  with qrySOrder do begin
      Close;
      SQL.Text:=' SELECT H141_EXAMPLCE 검사파트     '+#13#10+
                '      , H141_TSAMPLENO 샘플번호    '+#13#10+
                ' 　　 , H141_TAKEDAT 수거일자      '+#13#10+
                '      , H141_TAKETM 수거시간       '+#13#10+
                '      , H141_SEQNO 고유번호        '+#13#10+
                '      , H141_TAKESEQ 수거순번      '+#13#10+
                '      , H141_CHARTNO 차트번호      '+#13#10+
                '      , FN_PATIENT_INFO(H141_CHARTNO) 환자성명              '+#13#10+
                '      , FN_PATIENT_INFO(H141_CHARTNO, ''B'') 생년월일       '+#13#10+
                '      , FN_SEXAGE(H141_CHARTNO) 성별나이                    '+#13#10+
                '      , H141_VISTDAT 방문일자                               '+#13#10+
                '      , H141_ODRDAT 처방일자                                '+#13#10+
                '      , H141_ODRNO 처방번호                                 '+#13#10+
                '      , H141_ODRSEQ 처방서브번호                            '+#13#10+
                '      , H141_SUGACD 처방코드                                '+#13#10+
                '      , FN_SUGAMST_INFO( H141_SUGACD, ''H'') 한글명         '+#13#10+
                '      , FN_SUGAMST_INFO( H141_SUGACD, ''E'') 영문명         '+#13#10+
                '      , H141_RSLTYN 결과유무                                '+#13#10+
                '      , H141_NOTYYN 통보유무                                '+#13#10+
                '      , H141_SPECCD                                         '+#13#10+
                ' FROM TB_H141_LISTAKEBODY                                   '+#13#10+
                '    , TB_H131_SPPRESULT                                     '+#13#10+
                ' WHERE H141_TSAMPLENO = '''+TMaster.FBarCode+'''            '+#13#10+
                '   AND H141_TAKEDAT between '''+FormatDateTime('yyyymmdd', now-1)+'''  '+
                '                        And '''+FormatDateTime('yyyymmdd', now)+'''    '+#13#10+  // '수거일자 '
                '   AND NVL(H141_RSLTYN,'' '') IN(''N'', ''T'')                         '+#13#10+
                '   AND H141_SEQNO = H131_SEQNO                                         '+#13#10+
                '   AND (TRIM(H131_RESULT) IS NULL OR TRIM(H131_RESULT) = ''결과대기'') '+#13#10+
                '   AND H131_SPPTYPE = ''L010''                                         ';


                {' select O.*  '+#13#10+
                '      ,  FN_PATIENT_INFO(O.H130_CHARTNO) AS PNM '+#13#10+
                ' from           '+#13#10+
                ' TB_H130_SPPRECEIVE O     '+#13#10+
                ' Where 1=1               '+#13#10+
                '   And O.H130_TRANST <> ''F'' '+#13#10+
                '   And O.H130_SAMPLENO = '''+TMaster.FBarCode+'''  '+#13#10+
                '   AND O.H130_SPPTYPE = ''L010''         '+#13#10+
                '   AND O.H130_SEQNO between '''+FSEQ+''' and '''+TSEQ+''' ';}
      //ShowMessage(SQL.Text);


      try
          if SvrTEST = True then begin
              SQL.SaveToFile('.\오더조회.sql');

          end;

          Open;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= ' 오더체크 에러 입니다! 에러메세지->'+ e.Message;
              ShowMessage(e.Message);
              exit;
          end;
      end;

      if RecordCount > 0 then begin
          TMaster.vOrder:= VarArrayCreate([0, RecordCount-1], varVariant);
          TMaster.vAbbr := VarArrayCreate([0, RecordCount-1], varVariant);
          TMaster.vOrdList:= VarArrayCreate([0, RecordCount-1], varVariant);
          TMaster.vUpCode:=  VarArrayCreate([0, RecordCount-1], varVariant);
          TMaster.vOrdCdList:=  VarArrayCreate([0, RecordCount-1], varVariant);
          TMaster.vANO:=  VarArrayCreate([0, RecordCount-1], varVariant);
          K:=0;

          while Not Eof do begin
              //ECD:= Trim(FieldByName('H130_SUGACD').AsString);
              ECD:= Trim(FieldByName('처방코드').AsString);

              //ShowMessage(ECD);

              if TCode.SetCode_ECode(ECD) then begin
                  TMaster.vOrder[K]:= TCode.GetIfCode(ECD);
                  TMaster.vAbbr[K] := TCode.GetAbbr(ECD);
                  //TMaster.vOrdList[K]:= FieldByName('H130_SUGACD').AsString;
                  TMaster.vOrdList[K]:= FieldByName('처방코드').AsString;
                  TMaster.vUpCode[K]:= TCode.GetUpCode(ECD);
                  TMaster.vOrdCdList[K]:= '';
                  //TMaster.vANO[K]:= FieldByName('H130_SEQNO').AsString;
                  TMaster.vANO[K]:= FieldByName('고유번호').AsString;
                  Inc(K);

                  if Result = False then begin
                      TMaster.FOrdCode  := ''; //FieldByName('OCD').AsString;
                      //TMaster.FPID  := FieldByName('H130_CHARTNO').AsString;
                      //TMaster.FPNM  := FieldByName('PNM').AsString;
                      //TMaster.FANO  := FieldByName('H130_SEQNO').AsString;
                      TMaster.FPID  := FieldByName('차트번호').AsString;
                      TMaster.FPNM  := FieldByName('환자성명').AsString;
                      TMaster.FANO  := FieldByName('고유번호').AsString;
                      Result:= True;
                      TMaster.FOrdState:= 'Y';
                  end;
              end;

              Next;
          end;
      end;
  end;
end;

function TDM.UploadHosp_One_DJI(TMaster: TIfMaster): boolean;
begin
  Result:= False;

  if SvrConnection = False then begin
      TGlobal.ErrMsg:= 'Local 테스트중입니다!';
      exit;
  end;

  if TMaster.FANO = '' then exit;

  with qrySUp do begin
      Close;
      SQL.Text:=' select * from TB_H131_SPPRESULT  '+#13#10+
                '  where H131_SEQNO = '''+TMaster.FANO+'''            '+#13#10+
                '    and H131_SPPTYPE = ''L010''      ';
      if SvrTEST = True then
          SQL.SaveToFile('.\'+TMaster.FANO+'결과테이블조회.sql');

      Open;

      //NF00 수치타입,  CF00 문자타입
      if RecordCount > 0 then begin
          Close;
          SQL.Text:=' update TB_H131_SPPRESULT set H131_RESULT = '''+TMaster.FResult+'''              '+#13#10+
                    '                        , H131_SAVEDATE = To_Char(SysDate, ''YYYYMMDD'') '+#13#10+
                    '                        , H131_SAVETIME = To_Char(SysDate, ''HH24:MI:SS'') '+#13#10+
                    ' where H131_SEQNO = '''+TMaster.FANO+'''    '+#13#10+
                    '   And H131_SPPTYPE = ''L010''   ';
      end
      else begin
          Close;              //H131_CRUSERID, H131_CRUSERIP,  H131_RESULT2, H131_RESULT3, H131_DODOCT
          SQL.Text:=' Insert Into TB_H131_SPPRESULT (H131_SPPTYPE, H131_SEQNO, H131_RSLTFORM, H131_RESULT, H131_CRDTIME, H131_CRUSERID, '+
                    '                                H131_CRUSERIP, H131_UPDTIME, H131_UPUSERID, H131_UPUSERIP, H131_SAVEDATE, H131_SAVETIME, '+
                    '                                H131_SAVEMEN, H131_TRANSDATE, H131_TRANSTIME, H131_TRANSMEN, H131_ODRSERL, H131_RESULT2, H131_RESULT3, H131_DODOCT)'+
                    ' Select H130_SPPTYPE, H130_SEQNO, ''NF00'', '''+TMaster.FResult+''', H130_CRDTIME, H130_CRUSERID, '+
                    '        H130_CRUSERIP, H130_UPDTIME, H130_UPUSERID, H130_UPUSERIP, To_Char(SysDate, ''YYYYMMDD''), To_Char(SysDate, ''HH24:MI:SS''), '+
                    '        '''', '''', '''', '''', H130_ODRSERL, '''', '''', ''''  '+
                    ' From TB_H130_SPPRECEIVE                    '+#13#10+
                    ' where H130_SEQNO = '''+TMaster.FANO+'''    '+#13#10+
                    '   And H130_SPPTYPE = ''L010''   ';
      end;


      if SvrTEST = True then
          SQL.SaveToFile('.\'+TMaster.FANO+'결과등록.sql');

      try
          ExecSQL;
          Result:= True;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= 'SLA_LABRESULT 결과전송 에러 입니다! 에러메세지->'+ e.Message;
              ShowMessage(e.Message);
          end;
      end;
  end;
end;

function TDM.DownLoadOrder_DJI_WORK(FDT:string): string;
var
  K:integer;
  ECD:string;
  ODT:string;
  FSEQ, TSEQ:string;
  ICD, BCD, ANO, ADT, PID, SPC, PNM:string;
begin
  Result:= '';

  if SvrConnection = False then begin
      //TGlobal.ErrMsg:=  'Local TEST중입니다.';
      exit;
  end;

  ODT:= FDT;

  //인덱스 안타므로 SEQ를 미리 만들자..
  FSEQ:= ODT + '0000000';
  TSEQ:= ODT + '1000000';

  with qrySOrder do begin
      Close;
      SQL.Text:=' SELECT distinct H141_TSAMPLENO 샘플번호    '+#13#10+    //바코드
                ' 　　 , H141_TAKEDAT 수거일자      '+#13#10+
                '      , H141_TAKETM 수거시간       '+#13#10+
                //'      , H141_SEQNO 고유번호        '+#13#10+
                '      , H141_TAKESEQ 수거순번      '+#13#10+
                '      , H141_CHARTNO 차트번호      '+#13#10+
                '      , FN_PATIENT_INFO(H141_CHARTNO) 환자성명              '+#13#10+
                '      , FN_PATIENT_INFO(H141_CHARTNO, ''B'') 생년월일       '+#13#10+
                '      , FN_SEXAGE(H141_CHARTNO) 성별나이                    '+#13#10+
                '      , H141_VISTDAT 방문일자                               '+#13#10+
                '      , H141_ODRDAT 처방일자                                '+#13#10+
                '      , H141_ODRNO 처방번호                                 '+#13#10+
                '      , H141_ODRSEQ 처방서브번호                            '+#13#10+
                //'      , H141_SUGACD 처방코드                                '+#13#10+
                //'      , FN_SUGAMST_INFO( H141_SUGACD, ''H'') 한글명         '+#13#10+
                //'      , FN_SUGAMST_INFO( H141_SUGACD, ''E'') 영문명         '+#13#10+
                '      , H141_RSLTYN 결과유무                                '+#13#10+
                '      , H141_NOTYYN 통보유무                                '+#13#10+
                '      , H141_SPECCD                                         '+#13#10+
                ' FROM TB_H141_LISTAKEBODY                                   '+#13#10+
                '    , TB_H131_SPPRESULT                                     '+#13#10+
                ' WHERE H141_TAKEDAT between '''+FormatDateTime('yyyymmdd', now-1)+'''  '+
                '                        And '''+FormatDateTime('yyyymmdd', now)+'''    '+#13#10+  // '수거일자 '
                '   AND H141_SUGACD in '+TCode.FInQuery+#13#10+
                '   AND NVL(H141_RSLTYN,'' '') IN(''N'', ''T'')                         '+#13#10+
                '   AND H141_SEQNO = H131_SEQNO                                         '+#13#10+
                '   AND (TRIM(H131_RESULT) IS NULL OR TRIM(H131_RESULT) = ''결과대기'') '+#13#10+
                '   AND H131_SPPTYPE = '''+TGlobal.FSite+'''                            '+#13#10+       //L010: abga
                ' Order by 수거일자, 수거시간 ';

      try
          if SvrTEST = True then begin
              SQL.SaveToFile('.\WorkList.sql');

          end;

          Open;
      except
          on e:exception do
          begin
              TGlobal.ErrMsg:= ' 오더체크 에러 입니다! 에러메세지->'+ e.Message;
              ShowMessage(e.Message);
              exit;
          end;
      end;

      if RecordCount > 0 then begin
          while Not Eof do begin
              PID:= FieldByName('차트번호').AsString;
              PNM:= FieldByName('환자성명').AsString;
              SPC:= ''; //FieldByName('환자성명').AsString;
              BCD:= FieldByName('샘플번호').AsString;
              ADT:= FieldByName('수거일자').AsString;
              ANO:= FieldByName('수거순번').AsString;

              Result:= Result + PID + TAB + PNM + TAB + SPC + TAB + BCD + TAB + ADT + TAB + ANO + TAB + #13#10;

              Next;
          end;
      end;
  end;
end;

end.
