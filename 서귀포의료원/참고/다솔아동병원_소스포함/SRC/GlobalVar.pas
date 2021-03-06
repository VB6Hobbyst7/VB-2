unit GlobalVar;

interface

uses Forms, SysUtils, IniFiles, Windows, Variants, Dialogs, graphics,
     OleServer,
     JRO_TLB,
     ADODB,
     AdvGrid,StringLib, dateUtils;

const
  IniFileName = 'SANSOFT.Ini';
  ErrFileName = 'Error.Log';
  LogFileName = 'SVR.Log';
  DataFileName = 'RCV.Log';
  LoginXmlFileName = 'LOGIN.xml';
  DownXmlFileName = 'DOWN.xml';
  UpXmlFileName = 'UP.xml';
  InstXmlFileName = 'INST.xml';

  HospitalName = '';

  BarCodeLen = 12;
  PatIdLen = 8;

  const
   NULL = #0;
   SOH = #1;
   STX = #2;
   ETX = #3;
   EOT = #4;
   ENQ = #5;
   ACK = #6;
   TAB = #9;
   ETB = #23;
   LF  = #10;
   CR  = #13;
   DLE = #16;
   NAK = #21;
   FS  = #28;
   GS  = #29;
   RS  = #30;
   SHP = #35;
   TER = #60;
   HED = #62;
   EOS = #93;
   SYN = #$16;

type aArray = array of string;

  TGlobalVar = Class(TObject)
      Constructor Create;
    private
      FSvrMsg: string;
      FLoginMsg: string;
      FLogMsg: string;
      FDownMsg: string;
      FUpMsg: string;
    FIMsg: string;
    procedure SetLogMsg(cValue:string);
    procedure WriteLog(cFileName:string; cStr: string);
    procedure XmlLog(cFileName:string; cStr: string);
    procedure SetDataLog(const Value: string);
    procedure SetErrMsg(const Value: string);
    procedure SetSvrMsg(const Value: string);
    procedure SetDownMsg(const Value: string);
    procedure SetLoginMsg(const Value: string);
    procedure SetUpMsg(const Value: string);
    procedure SetIMsg(const Value: string);
    public
      AppPath:string;
      FICode, FIName, FTitle:string;
      LastMsg:string;
      UpState:integer;
      DpState:string;
      FErrMsg:string;
      FServerIP:string;
      FServerPt:integer;
      FSite:string;
      FUserID, FUserPwd, FUserNm:string;
      MainTop, MainLeft, MainHeight, MainWidth:integer;
      FAutoSend:boolean;
      property LogMsg:string read FLogMsg write SetLogMsg;
      property ErrMsg:string read FErrMsg write SetErrMsg;
      property SvrMsg:string read FSvrMsg write SetSvrMsg;
      property LoginMsg:string read FLoginMsg write SetLoginMsg;
      property DownMsg:string read FDownMsg write SetDownMsg;
      property UpMsg:string read FUpMsg write SetUpMsg;
      property IMsg:string read FIMsg write SetIMsg;

      property DataLog:string write SetDataLog;
      procedure LoadIni;
      procedure SaveIni;
      procedure MyFileCopy(FromFile, ToFile:string);
      procedure LocalMDBCompress(cDbName: string);
      function CompressAndRefair(cOldMdb, cNewMdb: string): boolean;
      function MousePosition(AppForm:TForm): TPoint;
  end;

type
  TResultMsg = record
    MsgNo: integer;
    Msg  : string;
  end;

var
  TGlobal: TGlobalVar;
  TuxMsg:string;
  
//procedure ShowMessage(const Str:string);
//function MessageDlg(const Msg: string; DlgType: TMsgDlgType;Buttons: TMsgDlgButtons; HelpCtx: Longint): Integer;
function Str2Double(cValue:string):double;
function Bool2Str(Value:boolean):string;
function Str2Bool(Value:string):boolean;
function Str2DateTime(Value:string):TDateTime;
function Str2ViewDate(Value:string):string;  //yyyymmdd -> yyyy-mm-dd
function Str2ViewTime(Value:string):string;  //hhnnss -> hh:nn:ss
function Str2ViewDTM(Value:string):string;   //yyyymmddhhnnss -> yyyy-mm-dd hh:nn:ss

function GetAgeSex(cAID:string):Variant;
function InsertRowIndex(TGrid:TAdvStringGrid):integer;
function GetAbbrIndex(var TGrid:TAdvStringGrid; Abbr:string):integer;
function GetGridDate(cDateTime:string):string;
procedure Delay(nTime: Cardinal);
function ASTMCheckSum(s:string):string;
function BCCCheckSum(cStr:string) : string;
function BCCCheckSum_Char(cStr:string) : Char;

function MakeStrArray(sData,sDilimeter:string):aArray;
function MakeSqlInStr(aInData:array of string):string;

function GetJusuData(BCD:string; var ADT, SLP, LNO:string):boolean;

implementation

uses Controls;
{
function MessageDlg(const Msg: string; DlgType: TMsgDlgType;Buttons: TMsgDlgButtons; HelpCtx: Longint): Integer;

var
  X1, Y1, X2, Y2:integer;
  X, Y:integer;
begin
  X1:= Screen.Forms[0].Left;
  X2:= Screen.Forms[0].Width;
  Y1:= Screen.ActiveForm.Top;
  Y2:= Screen.ActiveForm.Height;
  X:= (X2 - X1) div 2;
  Y:= (Y2 - Y1) div 2;

  Result:= MessageDlgPos(Msg, DlgType, Buttons, HelpCtx, X, Y);
end;

procedure ShowMessage(const Str:string);
var
  X1, Y1, X2, Y2:integer;
  X, Y:integer;
  Form:TForm;
  sName:string;
begin
  X1:= Screen.ActiveForm.Left;
  Y1:= Screen.ActiveForm.Top;
  X2:= Screen.ActiveForm.Width;
  Y2:= Screen.ActiveForm.Height;

  X:= (X2 - X1) div 2;
  Y:= (Y2 - Y1) div 2;

  ShowMessagePos(Str, X, Y);
end;
}

function GetJusuData(BCD:string; var ADT, SLP, LNO:string):boolean;
var
  cNow:string;
  nNow:integer;
  cDate:string;
  GetBarVal:integer;
  nAdd:integer;
  spcid:string;
begin
  Result:= False;

  ADT:=''; SLP:=''; LNO:= '';

  if Length(BCD) <> 12 then exit;

  nNow:= DaysBetween(Now, StrToDate('1999-01-01'));
  GetBarVal:= StrToInt(Copy(BCD,1,5));
  nAdd:= GetBarVal - nNow;

  ADT:= FormatDateTime('yyyymmdd', now+nAdd);
  SLP:= Copy(BCD, 6, 2);
  LNO:= Copy(BCD, 8, 5);

  Result:= True;
end;


function GetAgeSex(cAID:string):Variant;
var
  vTemp:Variant;
  cPatYear,cCurYear:string;
  cFlag:string;
  nAge:integer;
begin
  vTemp:=VarArrayCreate([0, 1], varVariant);
  vTemp[0]:='';
  vTemp[1]:='';

  GetAgeSex:=vTemp;

  if Length(cAID) < 13 then Exit;

  cPatYear:=Copy(cAID,1,2);
  cCurYear:=FormatDateTime('yyyy',Now);
  cFlag:=Copy(cAID,7,1);

  if cFlag[1] in ['3','4'] then
      nAge:=StrToInt(cCurYear)-StrToInt(cPatYear)-2000
  else
      nAge:=StrToInt(cCurYear)-StrToInt(cPatYear)-1900;

  vTemp[0]:=IntToStr(nAge);

  if cFlag[1] in ['1','3'] then
      vTemp[1]:='M'
  else
      vTemp[1]:='F';

  GetAgeSex:=vTemp;

end;

function Str2ViewDTM(Value:string):string;
begin
  Result:= '';
  if Length(Value) <> 14 then exit;

  Result:= Str2ViewDate(Copy(Value, 1, 8))+' '+Str2ViewTime(Copy(Value,9,6));
end;

function Str2ViewDate(Value:string):string;
begin
  Result:='';

  if Length(Value) = 6 then begin      //yymmdd
      Result:= Copy(Value,1,2) + '-' + Copy(Value, 3,2) + '-' + Copy(Value,5,2);
  end
  else
  if Length(Value) = 8 then begin      //yyyymmdd
      Result:= Copy(Value,1,4) + '-' + Copy(Value, 5,2) + '-' + Copy(Value,7,2);
  end
  else
      exit;
end;

function Str2ViewTime(Value:string):string;
begin
  Result:='';

  if Length(Value) = 4 then begin      //hhnn
      Result:= Copy(Value,1,2) + ':' + Copy(Value, 3,2);
  end
  else
  if Length(Value) = 6 then begin      //hhnnss
      Result:= Copy(Value,1,2) + ':' + Copy(Value, 3,2) + ':' + Copy(Value,5,2);
  end
  else
      exit;
end;

function MakeStrArray(sData,sDilimeter:string):aArray;
var
  nDil,i:integer;
  aStr:array of string;
begin
  Result:= nil;
  if Trim(sData) = '' then exit;
  if Length(sDilimeter) < 1 then exit;

  nDil:= CountStr(sData, sDilimeter)+1;
  sData:= sData+sDilimeter;
  SetLength(aStr, nDil);

  for i:= 0 to nDil -1 do begin
      aStr[i]:= Trim(TokenStr(sData,sDilimeter,i));
  end;

  Result:= aArray(aStr);
end;

function MakeSqlInStr(aInData:array of string):string;
var
  i:integer;
  sItem,sStr:string;
begin
  Result:='';

  for i:=Low(aInData) to High(aInData) do begin
      sItem:= Trim(aIndata[i]);
      if sItem = '' then continue;

      sStr:= sStr+',' + ''''+sItem+'''';
  end;

  Result:= Copy(sStr,2,Length(sStr)-1);
end;

function BCCCheckSum_Char(cStr:string) : Char;
var
  i, ll_asc : integer;
begin
    ll_asc:=0;

    for i := 1 to Length(cStr) do
    begin
        if i = 1 then
        begin
           ll_asc := ord(cStr[1]) xor ord(cStr[2]);
        end
        else if i = 2 then
        begin
           //1???? * 2???? ????
        end
        else
        begin
           ll_asc := ll_asc xor ord(cStr[i]);
        end;
    end;
  //  result := char(ll_asc);
    result := chr(ll_asc);

end;

function BCCCheckSum(cStr:string) : string;
var
  i, ll_asc : integer;
begin
    ll_asc:=0;
    
    for i := 1 to Length(cStr) do
    begin
        if i = 1 then
        begin
           ll_asc := ord(cStr[1]) xor ord(cStr[2]);
        end
        else if i = 2 then
        begin
           //1???? * 2???? ????
        end
        else
        begin
           ll_asc := ll_asc xor ord(cStr[i]);
        end;
    end;
  //  result := char(ll_asc);
    result := chr(ll_asc);
end;

function ASTMCheckSum(s:string):string;
var
   i, iSum : integer;
begin
   iSum := 0;
   for i := 1 to Length(s) do
      iSum := iSum + Ord(s[i]);
   result := StringReplace( Format('%4x', [iSum]) , ' ', '0', [rfReplaceAll]);
end;

procedure Delay(nTime: Cardinal);
var
  PastTime : dword;
begin
  PastTime := GetTickCount + (nTime);
  repeat
      Application.ProcessMessages;
  until GetTickCount > PastTime;
end;

function GetGridDate(cDateTime:string):string;
var
  yy,mm,dd:string;
begin
  Result:='';

  //yyyy-mm-dd hh:nn:ss
  if Length(cDateTime) < 10 then
      exit;

  yy:= Copy(cDateTime,1,4);
  mm:= Copy(cDateTime,6,2);
  dd:= Copy(cDateTime,9,2);

  Result:= yy+mm+dd;

end;

function GetAbbrIndex(var TGrid:TAdvStringGrid; Abbr:string):integer;
var
  i:integer;
begin
  Result:=0;
  for i:=1 to TGrid.AllColCount-1 do begin
      if Abbr = TGrid.Cells[i,0] then begin
          Result:= i;
          exit;
      end;
  end;

end;

function InsertRowIndex(TGrid:TAdvStringGrid):integer;
begin
  Result:= 1;
  TGrid.InsertRows(1,1);
end;

function Str2DateTime(Value:string):TDateTime;
var
  yy,mm,dd,hh,nn,ss:string;
  cTime:string;
begin
  //20080326112957
  yy:= Copy(Value,1,4);
  mm:= Copy(Value,5,2);
  dd:= Copy(Value,7,2);
  hh:= Copy(Value,9,2);
  nn:= Copy(Value,11,2);
  ss:= Copy(Value,13,2);

  cTime:= yy +'-'+ mm +'-'+ dd +' '+ hh +':'+ nn +':'+ ss;

  Result:= StrToDateTimeDef(cTime, Now);

end;

function Str2Bool(Value:string):boolean;
begin
  if UpperCase(Value) = 'True' then
      Result:= True
  else
      Result:= False;
end;

function Bool2Str(Value:boolean):string;
begin
  if Value then
      Result:= 'True'
  else
      Result:= 'False';
end;

function Str2Double(cValue:string):double;
begin
  Result:= StrToFloatDef(cValue, -100);
end;
{ TGlobalVar }

procedure TGlobalVar.MyFileCopy(FromFile, ToFile: string);
begin
  CopyFile(PChar(FromFile), PChar(ToFile), False);
end;

constructor TGlobalVar.Create;
begin
  AppPath:= ExtractFilePath(Application.ExeName);
  LoadIni;
end;

procedure TGlobalVar.LoadIni;
var
  TIni: TIniFile;
  i:integer;
  CTime:integer;
begin
  tIni:= TIniFile.Create(AppPath+ IniFileName);
  try
      with tIni do begin
          FServerIP := ReadString('HOSP', 'IP', '');
          FServerPt := ReadInteger('HOSP', 'PT', 80);
          FSite     := ReadString('HOSP', 'SITE', 'L010');
          FICode    := ReadString('INST', 'CODE', '');
          FIName    := ReadString('INST', 'NAME', '');
          FTITLE    := ReadString('INST', 'TITLE', '');
          FUserID   := ReadString('USER', 'ID', '');
          FAutoSend := ReadBool('HOSP', 'AUTO', False);

          MainTop   := ReadInteger('DISP', 'MTOP', 10);
          MainLeft  := ReadInteger('DISP', 'MLFT', 10);
          MainHeight  := ReadInteger('DISP', 'MHEI', 100);
          MainWidth  := ReadInteger('DISP', 'MWID', 100);
      end;

  finally
      tIni.Free;
  end;

end;

procedure TGlobalVar.SaveIni;
var
  TIni: TIniFile;
begin
  tIni:= TIniFile.Create(AppPath+ IniFileName);
  try
      with tIni do begin
          WriteString('HOSP', 'SITE', FSite);
          WriteString('USER', 'ID', FUserId);
          WriteBool('HOSP', 'AUTO', FAutoSend);

          WriteInteger('DISP', 'MTOP', MainTop);
          WriteInteger('DISP', 'MLFT', MainLeft);
          WriteInteger('DISP', 'MHEI', MainHeight);
          WriteInteger('DISP', 'MWID', MainWidth);
      end;

  finally
      tIni.Free;
  end;

end;

procedure TGlobalVar.SetLogMsg(cValue:string);
begin
  WriteLog(LoginXmlFileName, cValue);
end;

procedure TGlobalVar.WriteLog(cFileName:string; cStr: string);
var
  F:TextFile;
begin
try
  AssignFile(F,cFileName);

  if FileExists(cFileName) = False then
      Rewrite(F)
  else
      Append(F);

  try
      Writeln(F, FormatDateTime('yyyy-mm-dd hh:nn:ss', now) + #13#10 + cStr);

  finally
      CloseFile(F);
  end;
except
end;
end;

function TGlobalVar.CompressAndRefair(cOldMdb, cNewMdb: string): boolean;
var
  oJetEng: JetEngine;
  sOldMdb, sNewMdb: string;
begin
  Result:= False;

  sOldMdb:= 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+cOldMdb;
  sNewMdb:= 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+cNewMdb;

   try
     oJetEng := CoJetEngine.Create;
      oJetEng.CompactDatabase(sOldMdb, sNewMdb);
     oJetEng:= nil;
     Result:= True;
   except
     oJetEng:= nil;
     Result:= False;
   end;

end;

procedure TGlobalVar.LocalMDBCompress(cDbName: string);
var
  cNewMdb, cOldMdb: string;
  cFilePath: string;
begin
   try
     cFilePath:= ExtractFilePath(Application.ExeName);
     cOldMdb  := cFilePath + cDbName+'.MDB';
     cNewMdb  := cFilePath + cDbName+'_Back.MDB';

     if CompressAndRefair(cOldMdb, cNewMdb) then begin
       DeleteFile(Pchar(cOldMdb));
       RenameFile(cNewMdb, cOldMdb);
     end;
   except
   end;

end;

procedure TGlobalVar.SetDataLog(const Value: string);
begin
  WriteLog(DataFileName, Value);
end;

function TGlobalVar.MousePosition(AppForm:TForm): TPoint;
var
  TP:TPoint;
begin
  GetCursorPos(TP);
  TP:= AppForm.ScreenToClient(TP);
  Result:= TP;
end;

procedure TGlobalVar.SetErrMsg(const Value: string);
begin
  FErrMsg := Value;
  if FErrMsg <> '' then
      WriteLog(ErrFileName, Value);
end;

procedure TGlobalVar.SetSvrMsg(const Value: string);
begin
  FSvrMsg := Value;
  if FSvrMsg <> '' then
      WriteLog(LogFileName, Value);
end;

procedure TGlobalVar.XmlLog(cFileName, cStr: string);
var
  F:TextFile;
begin
try
  AssignFile(F,cFileName);

  //if FileExists(cFileName) = False then
      Rewrite(F);

  try
      Writeln(F, cStr);

  finally
      CloseFile(F);
  end;
except
end;

end;

procedure TGlobalVar.SetDownMsg(const Value: string);
begin
  FDownMsg := Value;
  XmlLog(DownXmlFileName, Value);
end;

procedure TGlobalVar.SetLoginMsg(const Value: string);
begin
  FLoginMsg := Value;
  XmlLog(LoginXmlFileName, Value);
end;

procedure TGlobalVar.SetUpMsg(const Value: string);
begin
  FUpMsg := Value;
  XmlLog(UpXmlFileName, Value);
end;

procedure TGlobalVar.SetIMsg(const Value: string);
begin
  FIMsg := Value;
  XmlLog(InstXmlFileName, Value);
end;

end.
