unit conTuxedo;

interface

uses SysUtils, Classes, Dialogs, Variants, TuxClient, GlobalVar;

type
  PTuxCaller = ^TuxCaller;

  TuxCaller = record
    Name:PChar;
    Data:string;
  end;

function TuxInit:boolean;
procedure TuxTerm;

//모듈분리 하려 했느나, 포인터 처리가 원하는대로 안된다..
function DownLoadTux_Str_TEST(BCD:string; var SvrMsg:string):boolean;

function DownLoadTux_Str(BCD:string; var SvrMsg:string):boolean;
function UpLoadTux_Str(BCD, ECD, RES:string; var SvrMsg:string):boolean;

function InitVar(Buffer:Pointer; var ErrMsg:string; GBN:integer=32):boolean;
function AddBufer_STR(Buffer:Pointer; FieldNm, Value:PChar; var ErrMsg:string; GBN:integer=32):boolean;
function AddBufer_INT(Buffer:Pointer; FieldNm:PChar; Value:integer; var ErrMsg:string; GBN:integer=32):boolean;
function TuxCall(SVCNM, InBuffer, OutBuffer:PChar; TransYN:boolean; var ErrMsg:string):integer;

implementation

uses stringlib;

function TuxInit:boolean;
var
  sFile, sLabel: string;
begin
  Result:= False;

  sFile := 'C:\cuh_95\wsenv.txt';      //화순:C:\cuh_tux\wsenv.txt
  sLabel := 'cuhocs';      //hcuhpmpa/

  if tuxreadenv(PChar(sFile), PChar(sLabel)) = -1 then
    ShowMessage(sFile+' read error ->' +StrPas(TPSTRERROR(GETTPERRNO)))
  else begin
    //ShowMessage('con tuxedo!');
  end;

  if tpinit(nil) = -1 then
  begin
    showmessage('tpinit failed --> '+ StrPas(TPSTRERROR(GETTPERRNO)));
    exit;
  end;

  Result:= True;
end;

procedure TuxTerm;
begin
  tpterm;
end;

function InitVar(Buffer:Pointer; var ErrMsg:string; GBN:integer):boolean;
begin
  Result:= False;

  if GBN = 16 then begin
      //FML

      Buffer:= tpalloc(FBufferType, '' , FBufferSize);
      if Buffer = nil then begin
          ErrMsg:= 'TPALLOC FAIL -> ' + StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;

      if Finit(Buffer, FBufferSize) = -1 then begin
          ErrMsg:= 'TPInit FAIL -> ' + StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;

      Result:= True;
  end
  else begin
      //FML32

      Buffer:= tpalloc(FBufferType32, '' , FBufferSize32);
      if Buffer = nil then begin
          ErrMsg:= 'TPALLOC FAIL -> ' + StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;

      if Finit32(Buffer, FBufferSize32) = -1 then begin
          ErrMsg:= 'TPInit FAIL -> ' + StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;
      Result:= True;
  end;
end;

function AddBufer_STR(Buffer:Pointer; FieldNm, Value:PChar; var ErrMsg:string; GBN:integer):boolean;
var
  FieldID:integer;
begin
  Result:= False;

  if GBN = 16 then begin
      //FML
      FieldID:= Fldid(FieldNm);
      if FieldID < 0 then begin
          ErrMsg:= StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;

      if Fchg(Buffer, FieldID, 0, Value, 0) < 0 then
      begin
          ErrMsg:= Fstrerror32(getFerror32());
          exit;
      end;

      Result:= True;
  end
  else begin
      //FML32
      FieldID:= Fldid32(FieldNm);
      if FieldID < 0 then begin
          ErrMsg:= StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;

      if Fchg32(Buffer, FieldID, 0, Value, 0) < 0 then
      begin
          ErrMsg:= Fstrerror32(getFerror32());
          exit;
      end;
      Result:= True;
  end;
end;

function AddBufer_INT(Buffer:Pointer; FieldNm:PChar; Value:integer; var ErrMsg:string;  GBN:integer):boolean;
var
  FieldID:integer;
begin
  Result:= False;

  if GBN = 16 then begin
      //FML
      FieldID:= Fldid(FieldNm);
      if FieldID < 0 then begin
          ErrMsg:= StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;

      if Fchg(Buffer, FieldID, 0, @Value, 0) < 0 then
      begin
          ErrMsg:= Fstrerror32(getFerror32());
          exit;
      end;

      Result:= True;
  end
  else begin
      //FML32
      FieldID:= Fldid32(FieldNm);
      if FieldID < 0 then begin
          ErrMsg:= StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;

      if Fchg32(Buffer, FieldID, 0, @Value, 0) < 0 then
      begin
          ErrMsg:= Fstrerror32(getFerror32());
          exit;
      end;
      Result:= True;
  end;

end;

function TuxCall(SVCNM, InBuffer, OutBuffer:PChar; TransYN:boolean; var ErrMsg:string):integer;
var
  Call, Len:integer;
begin
  Result:= -1;

  if TransYN then begin
      //For Transaction Mode
      if tpbegin(30,0) = -1 then
      begin
          ErrMsg:= 'TPBEGIN FAIL -> ' + StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;

      //Call
      Call:= tpcall(SVCNM, InBuffer, 0, @OutBuffer, @Len, 0);
      Result:= Call;

      if Call = -1 then begin
          //만약에 에러가 -> TPESYSTEM 라면 접속을 종료하고 재접속 하는걸 집어 넣어야 한다.
          ErrMsg:= 'tpcall failed --> '+ StrPas(TPSTRERROR(GETTPERRNO)) + '  PURCODE: ' + IntToStr(GETTPURCODE());
          tpabort(0);
          exit;
      end
      else begin
          if tpcommit(0) < 0 then begin
              ErrMsg:= 'TPCOMMIT FAIL -> ' + StrPas(TPSTRERROR(GETTPERRNO));
              tpabort(0);
              exit;
          end;
      end;
  end
  else begin
      //Call
      Call:= tpcall(SVCNM, InBuffer, 0, OutBuffer, @Len, 0);
      Result:= Call;

      if Call = -1 then begin
          ErrMsg:= 'tpcall failed --> '+ StrPas(TPSTRERROR(GETTPERRNO)) + '  PURCODE: ' + IntToStr(GETTPURCODE());
      end
  end;
end;

function DownLoadTux_Str_TEST(BCD:string; var SvrMsg:string):boolean;
var
  i, ii : integer;
  ADT, SNO, ICD: string;
  AID, SID, IID: string;
  PID, PNM, JNO: string;
  InBuffer, OutBuffer:PChar;
  FieldID:Integer;
  OutMsg:string;
begin
  Result:= False;
  SvrMsg:='';

  if Length(BCD) < 11 then exit;

  ADT:= '2'+Copy(BCD,1,7);
  SNO:= Copy(BCD,8,4);
  ICD:= TGlobal.FICode;

  AID:= 'acptacdt';
  SID:= 'acptsrno';
  IID:= 'pseudo89';

  InBuffer:= nil;
  OutBuffer:=nil;

  //Pointer형식이나 PChar 형식이던 var 형으로 Out이 되지 않는다..
  try  //~Finally
      //초기화
      if Not InitVar(InBuffer, SvrMsg, 16) then begin
          ShowMessage('InBuffer->'+SvrMsg); exit;
      end;

      if Not InitVar(OutBuffer, SvrMsg, 16) then begin
          ShowMessage('OutBuffer->'+SvrMsg); exit;
      end;

      //버퍼담기
      if Not AddBufer_STR(InBuffer, PChar(AID), PChar(ADT), SvrMsg, 16) then begin
          ShowMessage('Except: acptacdt -> '+ SvrMsg);
          exit;
      end;

      if Not AddBufer_STR(InBuffer, PChar(SID), PChar(SNO), SvrMsg, 16) then begin
          ShowMessage('Except: acptsrno -> '+ SvrMsg);
          exit;
      end;

      if Not AddBufer_STR(InBuffer, PChar(IID), PChar(ICD), SvrMsg, 16) then begin
          ShowMessage('Except: pseudo89 -> '+ SvrMsg);
          exit;
      end;

      //Call
      if TuxCall(PChar('OCP023LQ'), inBuffer, outBuffer, True, SvrMsg) > -1 then begin
          Result:= True;
          SvrMsg:='';

          //환자번호로 건수조회
          FieldID:= Fldid('pseudo70');

          ii:= Foccur(outBuffer, FieldID);

          //70:등록번호, 71:이름, 72+73:주민번호
          for i:=0 to ii -1 do begin
              PID:= StrPas(FVALS(outBuffer, Fldid('pseudo70') , i));
              PNM:= StrPas(FVALS(outBuffer, Fldid('pseudo71') , i));
              JNO:= StrPas(FVALS(outBuffer, Fldid('pseudo72') , i)) +
                    StrPas(FVALS(outBuffer, Fldid('pseudo73') , i));

              SvrMsg:= SvrMsg + STX + PID+'|'+PNM+'|'+JNO+ ETX;
          end;
      end
      else begin

      end;

  finally
      tpfree(InBuffer);
      tpfree(OutBuffer);
  end;
end;

function DownLoadTux_Str(BCD:string; var SvrMsg:string):boolean;
var
  i, ii, ix, Len : integer;
  ADT, SNO, ICD: string;
  AID, SID, IID: string;
  PID, PNM, JNO, ECD, ENM: string;
  InBuffer, OutBuffer:PChar;
  FieldID, nCnt:Integer;
begin
  Result:= False;
  SvrMsg:='';

  if Length(BCD) < 11 then exit;

  TuxInit;

  ADT:= '2'+Copy(BCD,1,7);
  SNO:= Copy(BCD,8,4);
  nCnt:= StrToInt(SNO);
  ICD:= TGlobal.FICode;

  AID:= 'acptacdt';
  SID:= 'acptsrno';
  IID:= 'pseudo89';

  InBuffer:= nil;
  OutBuffer:=nil;

  //Pointer형식이나 PChar 형식이던 var 형으로 Out이 되지 않는다..
  try  //~Finally
      //초기화
      InBuffer:= tpalloc(FBufferType, '' , FBufferSize);
      if InBuffer = nil then begin
          SvrMsg:= 'TPALLOC FAIL -> ' + StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;

      if Finit(InBuffer, FBufferSize) = -1 then begin
          SvrMsg:= 'TPInit FAIL -> ' + StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;

      OutBuffer:= tpalloc(FBufferType, '' , FBufferSize);
      if OutBuffer = nil then begin
          SvrMsg:= 'TPALLOC FAIL -> ' + StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;

      if Finit(OutBuffer, FBufferSize) = -1 then begin
          SvrMsg:= 'TPInit FAIL -> ' + StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;

      //버퍼담기
      //ADT
      FieldID:= Fldid(PChar(AID));
      if FieldID < 0 then begin
          SvrMsg:= StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;
      if Fchg(InBuffer, FieldID, 0, PChar(ADT), 0) < 0 then begin
          SvrMsg:= Fstrerror32(getFerror32());
          exit;
      end;

      //SPC
      FieldID:= Fldid(PChar(SID));
      if FieldID < 0 then begin
          SvrMsg:= StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;
      if Fchg(InBuffer, FieldID, 0, PChar(SNO), 0) < 0 then begin
          SvrMsg:= Fstrerror32(getFerror32());
          exit;
      end;

      //ICD
      FieldID:= Fldid(PChar(IID));
      if FieldID < 0 then begin
          SvrMsg:= StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;

      if Fchg(InBuffer, FieldID, 0, PChar(ICD), 0) < 0 then begin
          SvrMsg:= Fstrerror32(getFerror32());
          exit;
      end;

      if tpbegin(30,0) = -1 then
      begin
          SvrMsg:= 'TPBEGIN FAIL -> ' + StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;

      //Call          //OCP023LQ
      if tpcall(PChar('INF0104Q'), InBuffer, 0, @OutBuffer, @Len, 0) > -1 then begin
          if tpcommit(0) < 0 then begin
              ShowMessage('TPCOMMIT FAIL -> ' + StrPas(TPSTRERROR(GETTPERRNO)));
              tpabort(0);
              exit;
          end;
          
          Result:= True;
          SvrMsg:='';

          //환자번호로 건수조회
          FieldID:= Fldid('pseudo70');
          ii:= Foccur(outBuffer, FieldID);

          //검사코드조회
          FieldID:= Fldid('pseudo1');
          ix:= Foccur(outBuffer, FieldID);

          //70:등록번호, 71:이름, 72+73:주민번호
          for i:=0 to ii -1 do begin
              PID:= StrPas(FVALS(outBuffer, Fldid('pseudo70') , i));
              PNM:= StrPas(FVALS(outBuffer, Fldid('pseudo71') , i));
              JNO:= StrPas(FVALS(outBuffer, Fldid('pseudo72') , i))+
                    StrPas(FVALS(outBuffer, Fldid('pseudo73') , i));
          end;

          for i:=0 to ix -1 do begin
              ECD:= StrPas(FVALS(outBuffer, Fldid('pseudo1') , i));
              ENM:= StrPas(FVALS(outBuffer, Fldid('pseudo2') , i));
              SvrMsg:= SvrMsg + PID+'|'+PNM+'|'+JNO+'|'+ECD+'|'+ENM+'|'+ ETX;
          end;

          Result:= True;
      end
      else begin
          SvrMsg:= 'tpcall failed --> '+ StrPas(TPSTRERROR(GETTPERRNO)) + '  PURCODE: ' + IntToStr(GETTPURCODE());
          ShowMessage(SvrMsg);
      end;


  finally
      tpfree(InBuffer);
      tpfree(OutBuffer);
      TuxTerm;
  end;
end;

function UpLoadTux_Str(BCD, ECD, RES:string; var SvrMsg:string):boolean;
var
  i, ii, Len : integer;
  nCnt, FieldID:LongInt;
  ADT, SNO, ICD, STA, USR, CNT, Itm: string;
  AdtId, SnoId, IcdId, StaId, UsrId, CntId, EcdId, ItmId, ResID:string;
  InBuffer, OutBuffer:PChar;
begin
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

  Result:= False;
  SvrMsg:='';

  if Length(BCD) < 11 then exit;

  TuxInit;

  ADT:= '2'+Copy(BCD,1,7);
  SNO:= Copy(BCD,8,4);
  ICD:= TGlobal.FICode;
  USR:= TGlobal.FUserID;
  CNT:= '1';
  nCnt:= 1;
  STA:= 'T';
  ITM:= 'LIS';

  AdtId:= 'acptacdt';
  SnoId:= 'acptsrno';
  IcdId:= 'pseudo10';
  UsrId:= 'rslnuser';
  ItmId:= 'acptitem';
  EcdId:= 'rslnitem';
  CntId:= 'ocp_selcnt';
  ResId:= 'rslndscr';
  StaId:= 'rslnstat';

  InBuffer:= nil;
  OutBuffer:=nil;

  //Pointer형식이나 PChar 형식이던 var 형으로 Out이 되지 않는다..
  try  //~Finally
      //초기화
      InBuffer:= tpalloc(FBufferType, '' , FBufferSize);
      if InBuffer = nil then begin
          SvrMsg:= 'TPALLOC FAIL -> ' + StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;

      if Finit(InBuffer, FBufferSize) = -1 then begin
          SvrMsg:= 'TPInit FAIL -> ' + StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;

      OutBuffer:= tpalloc(FBufferType, '' , FBufferSize);
      if OutBuffer = nil then begin
          SvrMsg:= 'TPALLOC FAIL -> ' + StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;

      if Finit(OutBuffer, FBufferSize) = -1 then begin
          SvrMsg:= 'TPInit FAIL -> ' + StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;

      //버퍼담기
      //ADT
      FieldID:= Fldid(PChar(AdtId));
      if FieldID < 0 then begin
          SvrMsg:= StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;
      if Fchg(InBuffer, FieldID, 0, PChar(ADT), 0) < 0 then begin
          SvrMsg:= Fstrerror32(getFerror32());
          exit;
      end;

      //SNO
      FieldID:= Fldid(PChar(SnoId));
      if FieldID < 0 then begin
          SvrMsg:= StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;
      if Fchg(InBuffer, FieldID, 0, PChar(SNO), 0) < 0 then begin
          SvrMsg:= Fstrerror32(getFerror32());
          exit;
      end;

      //ICD
      FieldID:= Fldid(PChar(IcdId));
      if FieldID < 0 then begin
          SvrMsg:= StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;
      if Fchg(InBuffer, FieldID, 0, PChar(ICD), 0) < 0 then begin
          SvrMsg:= Fstrerror32(getFerror32());
          exit;
      end;

      //UsrId:= 'rslnuser';
      FieldID:= Fldid(PChar(UsrId));
      if FieldID < 0 then begin
          SvrMsg:= StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;
      if Fchg(InBuffer, FieldID, 0, PChar(USR), 0) < 0 then begin
          SvrMsg:= Fstrerror32(getFerror32());
          exit;
      end;

      //ItmId:= 'acptitem';
      FieldID:= Fldid(PChar(ItmId));
      if FieldID < 0 then begin
          SvrMsg:= StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;
      if Fchg(InBuffer, FieldID, 0, PChar(Itm), 0) < 0 then begin
          SvrMsg:= Fstrerror32(getFerror32());
          exit;
      end;

      //EcdId:= 'rslnitem';
      FieldID:= Fldid(PChar(EcdId));
      if FieldID < 0 then begin
          SvrMsg:= StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;
      if Fchg(InBuffer, FieldID, 0, PChar(ECD), 0) < 0 then begin
          SvrMsg:= Fstrerror32(getFerror32());
          exit;
      end;

      //CntId:= 'ocp_selcnt';
      FieldID:= Fldid(PChar(CntId));
      if FieldID < 0 then begin
          SvrMsg:= StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;
      if CFchg(InBuffer, FieldID, 0, @nCnt, 0, 1) < 0 then begin
          SvrMsg:= Fstrerror32(getFerror32());
          exit;
      end;

      //ResId:= 'rslndscr';
      FieldID:= Fldid(PChar(ResId));
      if FieldID < 0 then begin
          SvrMsg:= StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;
      if Fchg(InBuffer, FieldID, 0, PChar(RES), 0) < 0 then begin
          SvrMsg:= Fstrerror32(getFerror32());
          exit;
      end;

      //StaId:= 'rslndscr';
      FieldID:= Fldid(PChar(StaId));
      if FieldID < 0 then begin
          SvrMsg:= StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;
      if Fchg(InBuffer, FieldID, 0, PChar(STA), 0) < 0 then begin
          SvrMsg:= Fstrerror32(getFerror32());
          exit;
      end;

      if tpbegin(30,0) = -1 then
      begin
          SvrMsg:= 'TPBEGIN FAIL -> ' + StrPas(TPSTRERROR(GETTPERRNO));
          exit;
      end;

      //Call          //OCP034LA
      if tpcall(PChar('INF0104A'), InBuffer, 0, @OutBuffer, @Len, 0) > -1 then begin
          if tpcommit(0) < 0 then begin
              ShowMessage('TPCOMMIT FAIL -> ' + StrPas(TPSTRERROR(GETTPERRNO)));
              tpabort(0);
              exit;
          end;

          Result:= True;
          SvrMsg:='';
          {
          //환자번호로 건수조회
          FieldID:= Fldid('pseudo70');

          ii:= Foccur(outBuffer, FieldID);

          //70:등록번호, 71:이름, 72+73:주민번호
          for i:=0 to ii -1 do begin
              PID:= StrPas(FVALS(outBuffer, Fldid('pseudo70') , i));
              PNM:= StrPas(FVALS(outBuffer, Fldid('pseudo71') , i));
              JNO:= StrPas(FVALS(outBuffer, Fldid('pseudo72') , i))+
                    StrPas(FVALS(outBuffer, Fldid('pseudo73') , i))+
                    StrPas(FVALS(outBuffer, Fldid('rslnitem') , i))+
                    StrPas(FVALS(outBuffer, Fldid('itcditnm') , i));

              SvrMsg:= SvrMsg + STX + PID+'|'+PNM+'|'+JNO+ ETX;
          end;
          }
          Result:= True;
      end
      else begin
          SvrMsg:= 'tpcall failed --> '+ StrPas(TPSTRERROR(GETTPERRNO)) + '  PURCODE: ' + IntToStr(GETTPURCODE());
      end;


  finally
      tpfree(InBuffer);
      tpfree(OutBuffer);
      TuxTerm;
  end;

end;

end.
