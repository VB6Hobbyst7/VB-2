unit U_DM;

interface

uses
  SysUtils, Windows, Classes, DB, ADODB, Forms, U_Main, Dialogs, U_IFClass;

type
  TDPInfo = record
    FDelta,
    FPanic,
    FOutMsg:string;
    FRC:integer;
  end;

type
  TDM = class(TDataModule)
    procedure DataModuleCreate(Sender: TObject);
    function ServerConnection(var sMsg:string):boolean;
  private
  public
    procedure DeleteMaster(ExamDate, ExamSeq:string);
    procedure DeleteResult(ExamDate,ExamSeq:string);
    procedure DeleteOldData(nDays: Cardinal =60);
    function GetExamSeq(ExamDate:string):integer;
    function GetBarCodeSeq(ExamDate, BarCode:string):string;
    procedure DeleteData(ExamDate,ExamSeq:string);
    procedure GetExamData(IFCode, QCYN:string; var ExamCode, Abbr:string; var RefMin, RefMax:double); overload;
    procedure GetCodeData(ExamCode:string; var IfCode, Abbr:string; var RefMin, RefMax:double);
    function GetAbbr(ExamCode:string):string;
    function CheckSetCode(ExamCode:string):boolean;
    function GetIfCode(ExamCode:string):string;
    function GetExamCode(IfCode:string):string;
    function GetSelectOrder(BarCode, ExamCode:string):boolean;

    procedure UpdateSpcid(ExamDate, ExamSeq, NewSpcid:string);
    function GetQcBarCode(cLot, cIName:string):string;
    function SaveOneCode(Ecode,EName,Abbr,UpCode,SubCode,RefL,RefH,Seq:string):boolean;
    function DeleteOneCode(ExamCode:string):boolean;

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
    function MakeUploadData(ExamDate:string;
                            ExamSeq:integer;
                            var vSpcid:variant;
                            var vExamcode:variant;
                            var vResult:variant;
                            var vErrFlag:variant;
                            var vEquipcd:variant;
                            var vIUser:variant):integer;

    procedure SaveState(ExamDate, ExamSeq, State:string);

    procedure ChangeBarCode(ExamDate, ExamSeq, BarCode:string); overload;
    procedure ChangeBarCode(OldBcd, NewBcd:string); overload;
    function SavePatId(PatId, ExamDate:string; ExamSeq:integer):boolean;
    procedure DeleteOneData(ExamDate:string; ExamSeq:integer);

    procedure SaveLotNo(sLotNo:string);
    procedure ChangeLotNo(OldLot, NewLot:string);
    procedure DeleteLotNo(sLotNo:string);
    function UpdateLotNo(sOldLot, sNewLot:string):boolean;
    procedure DeleteQCCode(LotNo, ExamCode, UpCode:string);
    function SaveQCCode(sLot,sCode,sName,sAbbr,sUpCd:string;
    dLow,dHigh:double; iSeq:integer):boolean;
    function SelectLocalOrder(BarCode:string):string;
    function DownLoadOrder(var TMaster:TH7180If):boolean; overload;
    function DownLoadOrder_Result(var TMaster:TH7180If):boolean; overload;
    function DownLoadOrder(ExamDate, ExamSeq, BarCode:string;
                           var PID, PNM, AcptNo:string):boolean; overload;
    procedure SaveMaster(TMaster:TH7180If);
    procedure SaveOrderList(TMaster:TH7180If);
    procedure SaveResultList(TMaster:TH7180If);

    function DownAndSaveCode(ExamDate, ExamSeq, ExamCode:string; var IfCode:string):boolean; overload;
    procedure SaveResult(TMaster:TH7180If);
    function GetCheckLowHigh(sResVal:string; RefLow, RefHigh:double):string;
    function UpLoadResult(ExamDate, ExamSeq, BarCode:string; var SvrMsg:string):boolean; overload;
    function UpLoadResult(BarCode:string; var SvrMsg:string):boolean; overload;
    procedure ChangeState(ExamDate, ExamSeq, State:string); overload;
    procedure ChangeState(BarCode, State:string); overload;
    function GetSvrQcBarCode(s1,s2:string):string;
  end;

//const
  //WSDL_DLL_NAME = 'CALL_WSDL.dll';
  //GHCODE = 'C0501';  //??????????
  //GICODE = 'C1';  //??????????

var
  DM: TDM;
  HostConnectionYN:boolean;

implementation

uses SetDataBase, GlobalVar, Variants, StringLib, U_CodeInfo, U_Server;

{$R *.dfm}

  //function DownLoadOrder_BarCode_Dll(BarCode, ICode, KIGWAN_CODE:string):string; external WSDL_DLL_NAME;
  //function DownLoadOrder_WorkList_Dll(ICode, KIGWAN_CODE:string):string; external WSDL_DLL_NAME;
  //function UploadResult_Dll(BarCode, ICode, KIGWAN_CODE, vResult, vExamCode:Variant):integer; external WSDL_DLL_NAME;


procedure TDM.DataModuleCreate(Sender: TObject);
var
  sMsg:string;
begin

  DeleteOldData(30); //1????..
  TConnection.AllDisconect;
  TGlobal.LocalMDBCompress('SANSOFT');

  TGlobal.HostConnecting:= ServerConnection(sMsg);
  if Not TGlobal.HostConnecting then
      ShowMessage(sMsg);

end;

procedure TDM.DeleteData(ExamDate,ExamSeq:string);
begin
  DeleteResult(ExamDate,ExamSeq);
  DeleteMaster(ExamDate,ExamSeq);
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

function TDM.DeleteOneCode(ExamCode: string): boolean;
var
  TSql: TQueryInfo;
begin
  TSql:= TqueryInfo.Create;
  try
      with TSql do begin
          Clear;
          AddSql(' Delete From TB_Code ');
          AddSql(' Where ExamCode = '''+ExamCode+''' ');
          Result:= LocalExcute;
      end;

  finally
      TSql.Free;
  end;

end;

procedure TDM.DeleteMaster(ExamDate, ExamSeq:string);
var
  TSql: TQueryInfo;
begin
  TSql:= TQueryInfo.Create;
  try
      with TSql do begin
          AddSql(' Delete From TB_Master  ');
          AddSql(' Where ExamDate = '''+ExamDate+''' ');
          AddSql('   And ExamSeq  = '''+ExamSeq+'''  ');
          LocalExcute;
      end;

  finally
      TSql.Free;
  end;

end;

procedure TDM.DeleteResult(ExamDate, ExamSeq: string);
var
  TSql: TQueryInfo;
begin
  TSql:= TQueryInfo.Create;
  try
      with TSql do begin
          AddSql(' Delete From TB_Result  ');
          AddSql(' Where ExamDate = '''+ExamDate+''' ');
          AddSql('   And ExamSeq  = '''+ExamSeq+'''  ');
          LocalExcute;
      end;

  finally
      TSql.Free;
  end;

end;

function TDM.GetExamSeq(ExamDate: string):integer;
var
  TSql: TQueryInfo;
  QryEx: TAdoQuery;
begin
  Result:= 1;
  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Application);

  try
      with TSql do begin
        Clear;
        AddSql(' SELECT iif( isNull(Max(ExamSeq)), 1, Max(ExamSeq)+1) As MAXSEQ From TB_Master ');
        AddSql(' Where ExamDate = '''+ExamDate+'''         ');
        RCount:= LocalSelect(QryEx);

        if RCount > 0 then
            Result:= QryEx.Fields[0].AsInteger;
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

function TDM.SaveOneCode(Ecode,EName,Abbr,UpCode,SubCode,RefL,RefH,Seq:string): boolean;
var
  TSql: TQueryInfo;
  QryEx:TAdoQuery;
begin
  Result:= False;

  TSql:= TqueryInfo.Create;
  QryEx:= TADOQuery.Create(Application);

  try
      with TSql do begin
          Clear;
          AddSql(' Delete From TB_Code ');
          AddSql(' Where ExamCode = '''+ECode+''' ');
          LocalExcute;

          Clear;
          AddSql(' Insert Into TB_Code ');
          AddSql(' (ExamCode,ExamName, IFCode_Sub, IFCode, Abbr,RefLow,RefHigh,DispSeq) ');
          AddSql(' Values ');
          AddSql(' ('''+ECode+''', '''+EName+''', '''+SubCode+''', '''+UpCode+''', '''+Abbr+''' ');
          AddSql(' ,'+RefL+','+RefH+','+seq+' ) ');
          Result:= LocalExcute;
      end;

  finally
      QryEx.Free;
      TSql.Free;
  end;
end;

procedure TDM.UpdateSpcid(ExamDate, ExamSeq, NewSpcid:string);
var
  TSql: TQueryInfo;
begin
  TSql:= TQueryInfo.Create;
  try
      with TSql do
      begin
          Clear;
          AddSql(' Update TB_Master Set BarCode = '''+NewSpcid+''' ');
          AddSql(' Where ExamDate = '''+ExamDate+''' ');
          AddSql('   And ExamSeq  = '''+ExamSeq+''' ');
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
  strcmd:= 'net send ' + sIp + ' Error code->' + sState + '   SampleNo: ' +sSampleId;
  WinExec(pansichar(strcmd), SW_HIDE);
end;

procedure TDM.SendDeltaMessage(sIp, sState, sSampleId: string);
var
  strcmd:string;
begin
  strcmd:= 'net send ' + sIp + ' Error code-> ' + sState + '   SampleNo: ' +sSampleId + '[ Panic Data!!!! ]';
  WinExec(pansichar(strcmd), SW_HIDE);
end;

function TDM.MakeUploadData(ExamDate:string;
                            ExamSeq:integer;
                            var vSpcid:variant;
                            var vExamcode:variant;
                            var vResult:variant;
                            var vErrFlag:variant;
                            var vEquipcd:variant;
                            var vIUser:variant):integer;
var
  i,nOpNo:integer;
  TSql:TQueryInfo;
  QryEx:TAdoQuery;
  sOpId:string;
begin
  Result:=0;

  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Application);

  try
      with TSql do begin
          Clear;
          AddSql(' Select A.ExamDate, A.ExamSeq, A.Location, A.BarCode, A.PatNo, A.OpId ');
          AddSql('      , A.LotNo, A.QCYN, A.ICode, A.Flag, A.POCTYN          ');
          AddSql('      , B.IFCode, B.ExamCode, B.RsltTxt                  ');
          AddSql(' From TB_Master A Inner Join TB_Result B on (A.ExamDate=B.ExamDate And A.ExamSeq=B.ExamSeq) ');
          AddSql(' Where A.ExamDate = '''+ExamDate+''' ');
          AddSql('   And A.ExamSeq  = '+IntToStr(ExamSeq)+' ');
          AddSql(' Order By B.DispSeq ');
          RCount:= LocalSelect(QryEx);

          if RCount = 0 then
              exit;

          Result:= RCount;

          vSpcid    :=VarArrayCreate([0, RCount-1], VarOleStr);
          vExamcode :=VarArrayCreate([0, RCount-1], VarOleStr);
          vResult   :=VarArrayCreate([0, RCount-1], VarOleStr);
          vErrFlag  :=VarArrayCreate([0, RCount-1], VarOleStr);
          vEquipcd  :=VarArrayCreate([0, RCount-1], VarOleStr);
          vIUser    :=VarArrayCreate([0, RCount-1], VarOleStr);

          //VarClear(vSpcid);
          //VarClear(vExamcode);
          //VarClear(vResult);
          //VarClear(vErrFlag);
          //VarClear(vEquipcd);
          //VarClear(vIUser);

          sOpId:= QryEx.FieldByName('OPID').AsString;

          nOpNo:=StrToIntDef(Copy(sOpId,1,1),0);
          case nOpNo of           // ?????? ?????? ?? ?????? ???????? ?????? ???????? ????.
              1:sOpId:='N'+Trim(Copy(sOpId,2,5));
              2:sOpId:='R'+Trim(Copy(sOpId,2,5));
              3:sOpId:='X'+Trim(Copy(sOpId,2,5));
              4:sOpId:='Y'+Trim(Copy(sOpId,2,5));
              5:sOpId:='Z'+Trim(Copy(sOpId,2,5));
              6:sOpId:='E'+Trim(Copy(sOpId,2,5));
              7:sOpId:='G'+Trim(copy(sOpId,2,5));
              else
              sOpId:=sOpId;
          end;

          i:=0;
          with QryEx do begin
              while Not Eof do begin
                  vSpcid[i]   := FieldByName('BarCode').AsString;
                  vExamcode[i]:= FieldByName('ExamCode').AsString;
                  vResult[i]  := FieldByName('RsltTxt').AsString;
                  vErrFlag[i] := '0';
                  vEquipcd[i] := FieldByName('ICode').AsString;
                  vIUser[i]   := FieldByName('OpId').AsString;

                  inc(i);
                  Next;
              end;
          end;
      end;

  finally
      QryEx.Free;
      TSql.Free;
  end;
end;


procedure TDM.SaveState(ExamDate, ExamSeq, State:string);
var
  TSql:TQueryInfo;
begin
  TSql:= TQueryInfo.Create;
  try
      with TSql do begin
          Clear;
          SQLCmd:= ' Update TB_Master Set UpState = '''+State+''' '+
                   ' Where ExamDate = '''+ExamDate+'''            '+
                   '   And ExamSeq  = '''+ExamSeq+'''             ';
          LocalExcute;
      end;
  finally
      TSql.Free;
  end;
end;

function TDM.SavePatId(PatId, ExamDate: string; ExamSeq: integer):boolean;
var
  TSql:TQueryInfo;
begin
  Result:= False;
  if Trim(PatId) = '' then exit;

  TSql:= TQueryInfo.Create;

  try
      with TSql do begin
          Clear;
          AddSql(' Update TB_Master Set PatNo = '''+PatId+''' ');
          AddSql(' Where ExamDate = '''+ExamDate+''' ');
          AddSql('   And ExamSeq  = '+IntToStr(ExamSeq)+' ');
          Result:= LocalExcute;
      end;

  finally
      TSql.Free;
  end;

end;

procedure TDM.DeleteOneData(ExamDate: string; ExamSeq: integer);
var
  TSql:TQueryInfo;
begin
  TSql:= TQueryInfo.Create;
  try
      with TSql do begin
          Clear;
          AddSql(' Delete From TB_Result ');
          AddSql(' Where ExamDate = '''+ExamDate+''' ');
          AddSql('   And ExamSeq  = '+IntToStr(ExamSeq)+' ');
          LocalExcute;
          Clear;
          AddSql(' Delete From TB_Master ');
          AddSql(' Where ExamDate = '''+ExamDate+''' ');
          AddSql('   And ExamSeq  = '+IntToStr(ExamSeq)+' ');
          LocalExcute;
      end;
  finally
      TSql.Free;
  end;

end;

function TDM.ServerConnection(var sMsg:string):boolean;
var
  S:string;
begin
  Result:= False;

  //?????????? ????????.
  S:= WorkListCall;

  if S = '' then begin
      sMsg:= TGlobal.SvrError;
      exit;
  end
  else
      Result:= True;
end;

procedure TDM.SaveLotNo(sLotNo: string);
var
  i:integer;
  TSql: TQueryInfo;
  QryEx:TAdoQuery;
  sExamSeq,sDelta,sPanic,sCri:string;
  sExamCode:string;
begin
  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Application);

  try
      with TSql do begin
          Clear;
          SqlCmd:= ' Select LotNo From TB_QCM Where LotNo = '''+sLotNo+''' ';
          RCount:= LocalSelect(QryEx);
          if RCount > 0 then begin
              ShowMessage('???? ?????? Lot ??????!'); exit;
          end;

          SQLCmd:= ' Insert Into TB_QCM (LotNo) Values ('''+sLotNo+''') ';
          LocalExcute;
      end;

  finally
      Tsql.Free;
      QryEx.Free;
  end;

end;

function TDM.UpdateLotNo(sOldLot, sNewLot: string): boolean;
var
  i:integer;
  TSql: TQueryInfo;
  sExamSeq,sDelta,sPanic,sCri:string;
  sExamCode:string;
begin
  Result:= False;
  
  TSql:= TQueryInfo.Create;

  //Lot Master;
  try
      with TSql do begin
          Clear;
          AddSql(' Update TB_QCM Set LotNo = '''+sNewLot+''' Where LotNo = '''+sOldLot+''' ');
          LocalExcute;
      end;

  finally
      Tsql.Free;
  end;

end;

function TDM.SaveQCCode(sLot, sCode, sName, sAbbr, sUpCd: string; dLow,
  dHigh: double; iSeq: integer): boolean;
var
  i:integer;
  TSql: TQueryInfo;
  QryEx:TAdoQuery;
  sExamSeq,sDelta,sPanic,sCri:string;
  sExamCode:string;
begin
  Result:= False;

  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Self);

  try
      with TSql do begin
          Clear;
          AddSql(' Select LotNo From TB_QC ');
          AddSql(' Where LotNo = '''+sLot+''' ');
          AddSql('   And ExamCode = '''+sCode+''' ');
          AddSql('   And IFCode = '''+sUpCd+''' ');
          RCount:= LocalSelect(QryEx);

          if RCount > 0 then begin
              Clear;
              AddSql(' Update TB_QC Set ');
              AddSql('   ExamName = '''+sName+''' ');
              AddSql(' , Abbr = '''+sAbbr+''' ');
              AddSql(' , RefLow = '+FloatToStr(dLow) );
              AddSql(' , RefHigh= '+FloatToStr(dHigh) );
              AddSql(' , DispSeq= '+IntToStr(iSeq) );
              AddSql(' Where LotNo = '''+sLot+''' ');
              AddSql('   And ExamCode = '''+sCode+''' ');
              AddSql('   And IFCode = '''+sUpCd+''' ');
          end
          else begin
              Clear;
              AddSql(' Insert Into TB_QC (LotNo,ExamCode,IFCode,ExamName,Abbr,RefLow,RefHigh,DispSeq) ');
              AddSql(' Values ('''+sLot+''','''+sCode+''','''+sUpCd+''','''+sName+''','''+sAbbr+''', ');
              AddSql(' '+FloatToStr(dLow)+','+FloatToStr(dHigh)+','+IntToStr(iSeq)+') ');
          end;

          Result:= LocalExcute;
      end;

  finally
      Tsql.Free;
      QryEx.Free;
  end;

end;

procedure TDM.ChangeLotNo(OldLot, NewLot: string);
var
  i:integer;
  TSql: TQueryInfo;
  sExamSeq,sDelta,sPanic,sCri:string;
  sExamCode:string;
begin
  TSql:= TQueryInfo.Create;

  try
      with TSql do begin
          Clear;
          AddSql(' Update TB_QCM Set LotNo = '''+NewLot+''' Where LotNo = '''+OldLot+''' ');
          LocalExcute;
          Clear;
          AddSql( ' Update TB_QC Set LotNo = '''+NewLot+''' Where LotNo = '''+OldLot+''' ');
          LocalExcute;
      end;

  finally
      Tsql.Free;
  end;

end;

procedure TDM.DeleteLotNo(sLotNo: string);
var
  TSql:TQueryInfo;
begin
  TSql:= TQueryInfo.Create;
  try
      with tSql do begin
          SqlCmd:= ' Delete From TB_QCM Where LotNo = '''+sLotNo+''' ';
          LocalExcute;
          SQLCmd:= ' Delete From TB_QC Where LotNo = '''+sLotNo+''' ';
          LocalExcute;
      end;

  finally
      TSql.Free;
  end;

end;

procedure TDM.DeleteQCCode(LotNo, ExamCode, UpCode: string);
var
  TSql: TQueryInfo;
begin
  TSql:= TQueryInfo.Create;

  try
      with TSql do begin
          SQLCmd:=' Delete From TB_QC                 '+
                  ' Where LotNo    = '''+LotNo+'''    '+
                  '   And ExamCode = '''+ExamCode+''' '+
                  '   And IFCode   = '''+UpCode+'''   ';
          LocalExcute;
      end;

  finally
      Tsql.Free;
  end;


end;

procedure TDM.SaveMaster(TMaster:TH7180If);
var
  TSql: TQueryInfo;
  QryEx:TADOQuery;
begin

  TSql:= TQueryInfo.Create;
  QryEx:= TADOQuery.Create(Application);

  try
      with Tsql do begin
          with TMaster do begin
              SqlCmd:= ' Select * From TB_Master '+
                       ' Where BarCode = '''+TMaster.BarCode+''' ';
              RCount:= LocalSelect(QryEx);

              if RCount > 0 then begin
                  SqlCmd:= ' Update TB_Master Set '+
                           '   PatNo = '''+FPatId+''' '+
                           ' , PatNm = '''+FPatNm+''' '+
                           ' , ExamDate = '''+FExamDate+''' '+
                           ' Where BarCode = '''+TMaster.BarCode+''' ';
                  LocalExcute;
              end
              else begin
                  SQLCMD:= ' Insert Into TB_Master (BarCode ,PatNo, PatNm, ExamDate)  Values '+
                           ' ( '''+BarCode+''', '''+FPatId+''', '''+FPatNm+''','''+FExamDate+''')  ';
                  LocalExcute;
              end;
          end;
      end;
  finally
      QryEx.Free;
      Tsql.Free;
  end;

end;

procedure TDM.SaveResult(TMaster:TH7180If);
var
  TSql: TQueryInfo;
  QryEx:TAdoQuery;
begin
  if TMaster.FExamCode = '' then exit;

  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Self);
  try
      with TMaster do begin
          with Tsql do begin
              SQLCMD:= ' Select * From TB_Result  '+
                       ' Where ExamDate = '''+FExamDate+''' '+
                       '   And ExamSeq  = '''+FExamSeq+'''  '+
                       '   And IfCode   = '''+FIfCode+'''   ';

              if LocalSelect(QryEx) > 0 then begin
                  SQLCmd:= ' Update TB_Result Set RsltTxt = '''+FResult+''' '+
                           ' Where ExamDate = '''+FExamDate+''' '+
                           '   And ExamSeq  = '''+FExamSeq+'''  '+
                           '   And IFCode   = '''+FIfCode+'''   ';
                  LocalExcute;
              end
              else begin
                  SQLCmd:= ' Insert Into TB_Result (ExamDate, ExamSeq, IfCode, ExamCode, '+
                           ' RsltTxt) Values ('''+FExamDate+''','''+FExamSeq+''', '+
                           ' '''+FIfCode+''', '''+FExamCode+''', '''+FResult+''') ';
                  LocalExcute;
              end;
          end;
      end;
  finally
      Tsql.Free;
      QryEx.Free;
  end;

end;

procedure TDM.ChangeBarCode(ExamDate, ExamSeq, BarCode: string);
var
  TSql:TQueryInfo;
begin
  TSql:= TQueryInfo.Create;

  try
      with TSql do begin
          Clear;
          AddSql(' Update TB_Master Set BarCode = '''+BarCode+''' ');
          AddSql(' Where ExamDate = '''+ExamDate+''' ');
          AddSql('   And ExamSeq  = '''+ExamSeq+''' ');
          LocalExcute;
      end;

  finally
      TSql.Free;
  end;


end;

function TDM.GetCheckLowHigh(sResVal: string; RefLow, RefHigh: double): string;
var
  dVal:double;
begin
  Result:= '';

  dVal:= StrToFloatDef(sResVal, -1);
  if dVal < 0 then exit;

  if dVal < RefLow then begin
      Result:= 'L'; exit;
  end;

  if dVal > RefHigh then begin
      Result:= 'H'; exit;
  end;

end;

function TDM.UpLoadResult(ExamDate, ExamSeq, BarCode:string; var SvrMsg:string): boolean;
var
  TSql: TQueryInfo;
  QryEx:TAdoQuery;
  i,rc:integer;
  vExamCode, vResult:variant;
begin
  Result:= False;
  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Self);

  SvrMsg:= '';

  try
      with TSql do begin
          SQLCmD:= ' Select * From TB_Result '+
                   ' Where ExamDate = '''+ExamDate+''' '+
                   '   And ExamSeq  = '''+ExamSeq+'''  '+
                   '   And (RsltTxt <> '''' and RsltTxt is Not Null) '+
                   '   And (ExamCode <> '''' and ExamCode is Not Null) '+
                   '   And ORDYN = ''Y'' ';
          RCount:= LocalSelect(QryEx);

          if RCount = 0 then begin
              SvrMsg:= '?????????? ?????? ?????? ?????? ???? ??????!';
              exit;
          end;

          vExamCode:= VarArrayCreate([0,RCount-1], varOleStr);
          vResult  := VarArrayCreate([0,RCount-1], varOleStr);

          i:=0;
          with QryEx do begin
              while Not Eof do begin
                  vExamCode[i]:= FieldByName('ExamCode').AsString;
                  vResult[i]  := FieldByName('RsltTxt').AsString;
                  Inc(i);
                  Next;
              end;
          end;

          rc:= UploadCall(BarCode, vExamCode, vResult, SvrMsg);
          if Rc > 0 then begin
              Result:= True;
          end;
      end;
  finally
      Tsql.Free;
      QryEx.Free;
  end;
end;

procedure TDM.ChangeState(ExamDate, ExamSeq, State: string);
var
  TSql:TQueryInfo;
begin
  TSql:= TQueryInfo.Create;
  try
      with TSql do begin
          Clear;
          SQLCmd:= ' Update TB_Master Set UpState = '''+State+''' '+
                   ' Where ExamDate = '''+ExamDate+'''            '+
                   '   And ExamSeq  = '''+ExamSeq+'''             ';
          LocalExcute;
      end;
  finally
      TSql.Free;
  end;

end;

procedure TDM.GetCodeData(ExamCode: string; var IfCode, Abbr: string;
  var RefMin, RefMax: double);
var
  TSql: TQueryInfo;
  QryEx:TAdoQuery;
begin
  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Self);

  IfCode:=''; Abbr:= ''; RefMin:=0.0; RefMax:= 0.0;

  try
      with TSql do begin
          SQLCmd:= ' Select * From TB_CODE Where ExamCode = '''+ExamCode+''' ';
          RCount:= LocalSelect(QryEx);

          with QryEx do begin
              if RCount > 0 then begin
                  IfCode:= FieldByName('IfCode').AsString;
                  RefMin:= FieldByName('RefLow').AsFloat;
                  RefMax:= FieldByName('RefHigh').AsFloat;
                  Abbr  := FieldByName('Abbr').AsString;
              end;
          end;
      end;
  finally
      Tsql.Free;
      QryEx.Free;
  end;
end;

function TDM.DownLoadOrder(ExamDate, ExamSeq, BarCode: string; var PID,
  PNM, AcptNo: string): boolean;
var
  hRcv:string;
  ECode,IdTmp,NmTmp, IfCode:string;
  i:integer;
  Frame, TEMP:string;
  slList:TStringList;
begin
      {
      'MSH|^~\&|HL7|MMS|||20090309162034||ORU^R01|1a082e2:10e59b48c04:-2cf9:27695009|P|2.3||||||8859/1
      'PID|||200902120070^??????^400722^1^20090212^20090212^DefaultDomain^PI
      'PV1||E|C0501
      'OBR|1||||||20090309162034
      'OBX|1|ST|WC2420||||||||R'}
      
  Result:= False;
  AcptNo:='';
  PNm:= '';
  PID:='';

  hRcv:= OrderCall(BarCode);
  if hRcv = '' then exit;

  slList:= TStringList.Create;

  try
      slList.Delimiter:= #13;
      slList.DelimitedText:= hRcv;

      for i:=0 to slList.Count -1 do begin
          TEMP:= slList.Strings[i];
          if TEMP = '' then continue;

          Frame:= TokenStr(Temp, '|', 1);
          if Frame = 'PID' then begin
              NmTmp:= TokenStr(Temp, '^', 2);
              IdTmp:= TokenStr(Temp, '^', 3);
          end
          else
          if Frame = 'OBX' then begin
             ECode:= TokenStr(Temp, '|', 4);
             if DownAndSaveCode(ExamDate, ExamSeq, ECode, IfCode) then begin
                 if Result = False then begin
                     PId:= IdTmp;
                     PNm:= NmTmp;
                     Result:= True;
                 end;
             end;
          end;
      end;

  finally
      slList.Free;
  end;
end;

function TDM.CheckSetCode(ExamCode: string): boolean;
var
  TSql: TQueryInfo;
  QryEx:TAdoQuery;
begin
  Result:= False;

  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Self);

  try
      with TSql do begin
          SQLCmd:= ' Select * From TB_CODE Where ExamCode = '''+ExamCode+''' ';
          RCount:= LocalSelect(QryEx);

          if RCount > 0 then Result:= True;
      end;
  finally
      Tsql.Free;
      QryEx.Free;
  end;

end;

function TDM.GetSvrQcBarCode(s1, s2: string): string;
begin
  Result:= '';
end;

function TDM.DownAndSaveCode(ExamDate, ExamSeq, ExamCode: string; var IfCode:string):boolean;
var
  TSql:TQueryInfo;
  QryEx:TAdoQuery;
  sOldResult:string;
begin
  Result:= False;

  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Self);

  IfCode:= GetIfCode(ExamCode);
  if IfCode = '' then exit;

  try
      with TSql do begin
          Clear;
          SqlCmd:= ' Select * From TB_Result '+
                   ' Where ExamDate = '''+ExamDate+''' '+
                   '   And ExamSeq  = '''+ExamSeq+'''  '+
                   '   And ExamCode = '''+ExamCode+''' ';
          if LocalSelect(QryEx) > 0 then begin
              sOldResult:= Trim(QryEx.FieldByName('RsltTxt').AsString);

              SqlCmd:= ' Update TB_Result Set ORDYN = ''Y'' '+
                       ' Where ExamDate = '''+ExamDate+''' '+
                       '   And ExamSeq  = '''+ExamSeq+'''  '+
                       '   And ExamCode = '''+ExamCode+''' ';
              LocalExcute;
          end
          else begin
              SqlCmd:= ' Insert Into TB_Result (ExamDate, ExamSeq, ExamCode, IfCode, OrdYN) '+
                       ' Values ('''+ExamDate+''', '''+ExamSeq+''', '''+ExamCode+''', '''+IfCode+''', ''Y'' ) ';
              LocalExcute;
          end;
      end;

      Result:= True;

      {???? ?????? ???????? ?????? ???? ??????.}
      if sOldResult <> '' then
          IfCode:= '';

  finally
      Tsql.Free;
      QryEx.Free;
  end;
end;

function TDM.GetIfCode(ExamCode: string): string;
var
  TSql: TQueryInfo;
  QryEx:TAdoQuery;
begin
  Result:= '';

  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Self);

  try
      with TSql do begin
          SQLCmd:= ' Select IfCode From TB_CODE Where ExamCode = '''+ExamCode+''' ';
          RCount:= LocalSelect(QryEx);

          if RCount > 0 then
              Result:= QryEx.Fields[0].AsString;
      end;
  finally
      Tsql.Free;
      QryEx.Free;
  end;

end;

function TDM.DownLoadOrder(var TMaster: TH7180If): boolean;
var
  hRcv:string;
  ECode,PNm, PId, IfCode:string;
  i:integer;
  Frame, TEMP:string;
  slList:TStringList;
begin
      {
      'MSH|^~\&|HL7|MMS|||20090309162034||ORU^R01|1a082e2:10e59b48c04:-2cf9:27695009|P|2.3||||||8859/1
      'PID|||200902120070^??????^400722^1^20090212^20090212^DefaultDomain^PI
      'PV1||E|C0501
      'OBR|1||||||20090309162034
      'OBX|1|ST|WC2420||||||||R'}

  TMaster.slIfCode.Clear;
  TMaster.slExCode.Clear;
  TMaster.slIfCode_Down.Clear;

  Result:= False;

  hRcv:= OrderCall(TMaster.BarCode);
  if hRcv = '' then exit;

  slList:= TStringList.Create;
  try
      slList.Delimiter:= #13;
      slList.DelimitedText:= hRcv;

      for i:=0 to slList.Count -1 do begin
          TEMP:= slList.Strings[i];
          if TEMP = '' then continue;

          Frame:= TokenStr(Temp, '|', 1);
          if Frame = 'PID' then begin
              PNm:= TokenStr(Temp, '^', 2);
              PId:= TokenStr(Temp, '^', 3);
          end
          else
          if Frame = 'OBX' then begin
             ECode:= TokenStr(Temp, '|', 4);
             TMaster.slExCode.Add(ECode);

             IfCode:= GetIfCode(ECode);
             TMaster.slIfCode_Down.Add(IfCode);

             if Result = False then begin
                 TMaster.FPatId:= PId;
                 TMaster.FPatNm:= PNm;
                 TMaster.FAcptNo:= '';
             end;

             if DM.GetSelectOrder(TMaster.BarCode, ECode) then begin
                 if IfCode <> '' then begin
                     TMaster.slIfCode.Add(IfCode);
                     Result:= True;
                     TMaster.FOrdState:= 'Y';
                 end;
             end;
          end;
      end;

  finally
      slList.Free;
  end;
end;

function TDM.GetBarCodeSeq(ExamDate, BarCode:string): string;
var
  TSql: TQueryInfo;
  QryEx: TAdoQuery;
begin
  Result:= '001';

  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Application);

  try
      with TSql do begin
        Clear;
        AddSql(' SELECT ExamSeq From TB_Master ');
        AddSql(' Where ExamDate = '''+ExamDate+''' ');
        AddSql('   And BarCode  = '''+BarCode+''' ');
        AddSql(' Order By ExamTime Desc ');
        RCount:= LocalSelect(QryEx);

        if RCount > 0 then
            Result:= QryEx.Fields[0].AsString
        else
            Result:= PadLeftStr(IntToStr(GetExamSeq(ExamDate)), '0', 3);
    end;

  finally
      TSql.Free;
      QryEx.Free;
  end;
end;

procedure TDM.GetExamData(IFCode, QCYN:string; var ExamCode, Abbr:string; var RefMin, RefMax:double);
var
  TSql: TQueryInfo;
  QryEx:TAdoQuery;
begin
  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Self);

  ExamCode:=''; Abbr:= ''; RefMin:=0.0; RefMax:= 0.0;

  try
      with TSql do begin
          if QCYN = 'Y' then
              SQLCmD:= ' Select * From TB_CODE         '+
                       ' Where IFCode_Sub = '''+IFCode+''' '
          else
              SQLCmD:= ' Select * From TB_CODE         '+
                       ' Where IFCode = '''+IFCode+''' ';

          RCount:= LocalSelect(QryEx);

          with QryEx do begin
              if RCount > 0 then begin
                  ExamCode:= FieldByName('ExamCode').AsString;
                  RefMin  := FieldByName('RefLow').AsFloat;
                  RefMax  := FieldByName('RefHigh').AsFloat;
                  Abbr    := FieldByName('Abbr').AsString;
              end;
          end;
      end;
  finally
      Tsql.Free;
      QryEx.Free;
  end;

end;

function TDM.SelectLocalOrder(BarCode: string): string;
var
  TSql:TQueryInfo;
  QryEx:TAdoQuery;
  sOldResult:string;
  IfCode:string;
begin
  Result:= '';

  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Self);

  try
      with TSql do begin
          Clear;
          SqlCmd:= ' Select * From TB_Result         '+
                   ' Where BarCode = '''+BarCode+''' ';

          if LocalSelect(QryEx) = 0 then exit;

          while Not QryEx.Eof do begin
              Result:= Result + QryEx.FieldByName('IfCode').AsString + '|' ;
              QryEx.Next;
          end;
      end;

  finally
      Tsql.Free;
      QryEx.Free;
  end;
end;

function TDM.GetSelectOrder(BarCode, ExamCode: string): boolean;
var
  TSql:TQueryInfo;
  QryEx:TAdoQuery;
  sOldResult:string;
begin
  Result:= False;

  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Self);

  try
      with TSql do begin
          Clear;
          SqlCmd:= ' Select * From TB_Result '+
                   ' Where BarCode  = '''+BarCode+''' '+
                   '   And ExamCode = '''+ExamCode+''' ';
                   //'   And ORDYN = ''Y'' ';

          if LocalSelect(QryEx) > 0 then begin
              sOldResult:= Trim(QryEx.FieldByName('RsltTxt').AsString);
              if sOldResult = '' then
                  Result:= True
              else
                  exit;
          end
          else
              Result:= True;
      end;

  finally
      Tsql.Free;
      QryEx.Free;
  end;
end;

function TDM.GetAbbr(ExamCode: string): string;
var
  TSql: TQueryInfo;
  QryEx:TAdoQuery;
begin
  Result:= '';

  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Self);

  try
      with TSql do begin
          SQLCmd:= ' Select * From TB_CODE Where ExamCode = '''+ExamCode+''' ';
          RCount:= LocalSelect(QryEx);

          with QryEx do begin
              if RCount > 0 then begin
                  Result  := FieldByName('Abbr').AsString;
              end;
          end;
      end;
  finally
      Tsql.Free;
      QryEx.Free;
  end;

end;

procedure TDM.SaveOrderList(TMaster: TH7180If);
var
  TSql: TQueryInfo;
  QryEx:TADOQuery;
  i:integer;
begin

  TSql:= TQueryInfo.Create;
  QryEx:= TADOQuery.Create(Application);

  try

      for i:=0 to TMaster.slExCode.Count -1 do begin
          Tsql.SqlCmd:= ' Select * From TB_Result '+
                        ' Where BarCode = '''+TMaster.BarCode+''' '+
                        '   And ExamCode = '''+TMaster.slExCode.Strings[i]+''' ';
          Tsql.RCount:= Tsql.LocalSelect(QryEx);

          if TSql.RCount > 0 then begin
              TSql.SqlCmd:= ' Update TB_Result Set  ORDYN = ''Y'''+
                            ' Where BarCode = '''+TMaster.BarCode+''' '+
                            '   And ExamCode = '''+TMaster.slExCode.Strings[i]+''' ';
              TSql.LocalExcute;
          end
          else begin
              TSql.SqlCmd:= ' Insert Into TB_Result (BarCode, ExamDate, ExamCode, IfCode, ORDYN) Values '+
                            ' ('''+TMaster.BarCode+''', '+
                            '  '''+TMaster.FExamDate+''', '+
                            '  '''+TMaster.slExCode.Strings[i]+''', '+
                            '  '''+TMaster.slIfCode_Down.Strings[i]+''', ''Y'' ) ';
              TSql.LocalExcute;
          end;
      end;

  finally
      QryEx.Free;
      Tsql.Free;
  end;

end;

procedure TDM.SaveResultList(TMaster: TH7180If);
var
  TSql: TQueryInfo;
  QryEx:TADOQuery;
  i:integer;
begin

  TSql:= TQueryInfo.Create;
  QryEx:= TADOQuery.Create(Application);

  try

      for i:=0 to TMaster.slResExCode.Count -1 do begin
          Tsql.SqlCmd:= ' Select * From TB_Result '+
                        ' Where BarCode  = '''+TMaster.BarCode+''' '+
                        '   And ExamCode = '''+TMaster.slResExCode.Strings[i]+''' ';
          Tsql.RCount:= Tsql.LocalSelect(QryEx);
          if TSql.RCount > 0 then begin
              TSql.SqlCmd:= ' Update TB_Result Set RsltTxt = '''+TMaster.slResult.Strings[i]+''' '+
                            ' Where BarCode  = '''+TMaster.BarCode+''' '+
                            '   And ExamCode = '''+TMaster.slResExCode.Strings[i]+''' ';
              TSql.LocalExcute;
          end
          else begin
              TSql.SqlCmd:= ' Insert Into TB_Result (BarCode, ExamDate, ExamCode, IfCode, RsltTxt) Values '+
                            ' ('''+TMaster.BarCode+''', '+
                            '  '''+TMaster.FExamDate+''', '+
                            '  '''+TMaster.slResExCode.Strings[i]+''', '+
                            '  '''+TMaster.slResIfCode.Strings[i]+''', '+
                            '  '''+TMaster.slResult.Strings[i]+''' ) ';
              TSql.LocalExcute;
          end;
      end;

  finally
      QryEx.Free;
      Tsql.Free;
  end;

end;

function TDM.GetExamCode(IfCode: string): string;
var
  TSql: TQueryInfo;
  QryEx:TAdoQuery;
begin
  Result:= '';

  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Self);

  try
      with TSql do begin
          SQLCmd:= ' Select ExamCode From TB_CODE Where IfCode = '''+IfCode+''' ';
          RCount:= LocalSelect(QryEx);

          if RCount > 0 then
              Result:= QryEx.Fields[0].AsString;
      end;
  finally
      Tsql.Free;
      QryEx.Free;
  end;

end;

function TDM.UpLoadResult(BarCode: string; var SvrMsg: string): boolean;
var
  TSql: TQueryInfo;
  QryEx:TAdoQuery;
  i,rc:integer;
  vExamCode, vResult:variant;
begin
  Result:= False;
  TSql:= TQueryInfo.Create;
  QryEx:= TAdoQuery.Create(Self);

  SvrMsg:= '';

  try
      with TSql do begin
          SQLCmD:= ' Select * From TB_Result '+
                   ' Where BarCode = '''+BarCode+''' '+
                   '   And (RsltTxt <> '''' and RsltTxt is Not Null) '+
                   '   And (ExamCode <> '''' and ExamCode is Not Null) '+
                   '   And ORDYN = ''Y'' ';
          RCount:= LocalSelect(QryEx);

          if RCount = 0 then begin
              SvrMsg:= '?????????? ?????? ?????? ?????? ???? ??????!';
              exit;
          end;

          vExamCode:= VarArrayCreate([0,RCount-1], varOleStr);
          vResult  := VarArrayCreate([0,RCount-1], varOleStr);

          i:=0;
          with QryEx do begin
              while Not Eof do begin
                  vExamCode[i]:= FieldByName('ExamCode').AsString;
                  vResult[i]  := FieldByName('RsltTxt').AsString;
                  Inc(i);
                  Next;
              end;
          end;

          rc:= UploadCall(BarCode, vExamCode, vResult, SvrMsg);
          if Rc > 0 then begin
              Result:= True;
          end;
      end;
  finally
      Tsql.Free;
      QryEx.Free;
  end;
end;

procedure TDM.ChangeState(BarCode, State: string);
var
  TSql:TQueryInfo;
begin
  TSql:= TQueryInfo.Create;
  try
      with TSql do begin
          Clear;
          SQLCmd:= ' Update TB_Master Set UpState = '''+State+''' '+
                   ' Where BarCode = '''+BarCode+'''            ';
          LocalExcute;
      end;
  finally
      TSql.Free;
  end;

end;

function TDM.DownLoadOrder_Result(var TMaster: TH7180If): boolean;
var
  hRcv:string;
  ECode,PNm, PId, IfCode:string;
  i:integer;
  Frame, TEMP:string;
  slList:TStringList;
begin
      {
      'MSH|^~\&|HL7|MMS|||20090309162034||ORU^R01|1a082e2:10e59b48c04:-2cf9:27695009|P|2.3||||||8859/1
      'PID|||200902120070^??????^400722^1^20090212^20090212^DefaultDomain^PI
      'PV1||E|C0501
      'OBR|1||||||20090309162034
      'OBX|1|ST|WC2420||||||||R'}

  TMaster.slIfCode.Clear;
  TMaster.slExCode.Clear;
  TMaster.slIfCode_Down.Clear;

  Result:= False;

  hRcv:= OrderCall(TMaster.BarCode);
  if hRcv = '' then exit;

  slList:= TStringList.Create;
  try
      slList.Delimiter:= #13;
      slList.DelimitedText:= hRcv;

      for i:=0 to slList.Count -1 do begin
          TEMP:= slList.Strings[i];
          if TEMP = '' then continue;

          Frame:= TokenStr(Temp, '|', 1);
          if Frame = 'PID' then begin
              PNm:= TokenStr(Temp, '^', 2);
              PId:= TokenStr(Temp, '^', 3);
          end
          else
          if Frame = 'OBX' then begin
             ECode:= TokenStr(Temp, '|', 4);
             TMaster.slExCode.Add(ECode);

             IfCode:= GetIfCode(ECode);
             TMaster.slIfCode_Down.Add(IfCode);

             if Result = False then begin
                 TMaster.FPatId:= PId;
                 TMaster.FPatNm:= PNm;
                 TMaster.FAcptNo:= '';
             end;

             TMaster.FOrdState:= 'Y';
             Result:= True;

          end;
      end;

      SaveOrderList(TMaster);

  finally
      slList.Free;
  end;
end;

procedure TDM.ChangeBarCode(OldBcd, NewBcd: string);
var
  TSql:TQueryInfo;
begin
  TSql:= TQueryInfo.Create;

  try
      with TSql do begin
          Clear;
          AddSql(' Update TB_Master Set BarCode = '''+NewBcd+''' ');
          AddSql(' Where BarCode = '''+OldBcd+''' ');
          LocalExcute;

          Clear;
          AddSql(' Update TB_Result Set BarCode = '''+NewBcd+''' ');
          AddSql(' Where BarCode = '''+OldBcd+''' ');
          LocalExcute;
      end;

  finally
      TSql.Free;
  end;

end;

end.
