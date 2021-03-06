unit GlobalVar;

interface

uses Forms, SysUtils, IniFiles, Windows, Variants, Dialogs, graphics,
     OleServer,
     JRO_TLB,
     ADODB,
     AdvGrid,
     Classes;

const
  IniFileName = 'SANSOFT.Ini';
  LogFileName = 'Error.Log';
  DataFileName = 'RCV.Log';

  RCVFileName = 'H7180RCV.LOG';
  SENDFileName = 'SEND.LOG';

  EnvFile = 'sl.env';
  Color_Low = clRed;
  Color_High = clRed;

  USE_THREAD = False;

  BarCodeLen = 12;
  PatNoLen = 8;

  AppName = 'ABL';

  DEFTOP = 100;
  DEFLFT = 100;
  DEFWID = 947;
  DEFHEI = 743;

  const
   NULL = #0;
   SOH = #1;
   STX = #2;
   ETX = #3;
   EOT = #4;
   ENQ = #5;
   ACK = #6;
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

   // Hitachi 7180
   ENDFRAME = ':';
   FR1FRAME = '1';
   FR2FRAME = '2';
   FR9FRAME = '9';
   SPEFRAME = ';';
   SPMFRAME = '<';
   OPCFRAME = '=';
   ANYFRAME = '>';
   REPFRAME = '?';
   SUSFRAME = '@';
   RECFRAME = 'A';

   MOR = '>' + #3 + '3E';     // More
   //REP = '?' + #3 + '3F';     // Repeat

type
  OS_Lang = (KOR=1042, ENG=1033);
  PlayMode = (Releas, Debug);
  TCommType = (ctMSCOMM, ctCPort, ctTMSComm);

type
  RCommSet = Record
    BaudRate,
    Parity,
    DataBit,
    StopBit:string;
    PortNum:integer;
    HandShake:integer;
    Dtr, Rts:boolean;
    Settings:string;
  end;

type
  TGlobalVar = Class(TObject)
      ComPortSet: RCommSet;
      AppPath:string;
      AppTitle:string;
      OsLang:OS_Lang;
      LastMsg:string;
      FInstName:string;
      FUserId:string;
      IPAddr:string;
      Constructor Create;
      destructor Destroy; override;
    private
      FLogMsg:string;
      procedure SetLogMsg(cValue:string);
      procedure WriteLog(cFileName:string; cStr: string);
      procedure SetDataLog(const Value: string);
    procedure SetSaveLog(const Value: string);
    public
      HostConnecting:boolean;
      SvrError:string;
      MainTop,
      MainLeft,
      MainWidth,
      MainHeigh:integer;
      property LogMsg:string read FLogMsg write SetLogMsg;
      procedure LoadIni;
      procedure SaveIni;
      procedure ClearLog;
      procedure ComPortIniLoad;
      procedure ComPortIniSave;
      procedure MyFileCopy(FromFile, ToFile:string);
      procedure LocalMDBCompress(cDbName: string);
      function CompressAndRefair(cOldMdb, cNewMdb: string): boolean;
      property DataLog:string write SetDataLog;
      property SavePacket:string write SetSaveLog;
      function GetIpAddr:string;
  end;

var
  TGlobal: TGlobalVar;

function Str2Double(cValue:string):double;
function Bool2Int(Value:boolean):integer;
function Int2Bool(Value:integer):boolean;
function Bool2Str(Value:boolean):string;
function Str2Bool(Value:string):boolean;
function Str2DateTime(Value:string):TDateTime;

function AddRowIndex(var TGrid:TAdvStringGrid):integer;
function GetAbbrIndex(var TGrid:TAdvStringGrid; Abbr:string):integer;
function GetGridDate(cDateTime:string):string;
function ViewDateTime(ExamDateTime:string):string;
procedure Delay(nTime: Cardinal);
function GetSampleDateTime(cData:string):TDateTime;

function CheckBoxCheckYN(Grid:TAdvStringGrid): boolean;

implementation
uses WinSock;

function Bool2Int(Value:boolean):integer;
begin
  if Value then
      Result:=1
  else
      Result:=0;
end;

function Int2Bool(Value:integer):boolean;
begin
  if Value = 0 then
      Result:= False
  else
      Result:= True;
end;

function CheckBoxCheckYN(Grid:TAdvStringGrid): boolean;
var
  i:integer;
  bCheck:boolean;
begin
  Result:= False;
  for i:=1 to Grid.RowCount -1 do begin
      bCheck:= False;
      Grid.GetCheckBoxState(0, i, bCheck);
      if bCheck then begin
          Result:= True; exit;
      end;
  end;

  ShowMessage('?????? ?????? ????????!');

end;


function ViewDateTime(ExamDateTime:string):string;
var
  Len:integer;
  DT:string;
begin
  Len:= Length(ExamDateTime);
  DT:= ExamDateTime;

  Case Len of
      8: Result:= Copy(DT,1,4) + '-' + Copy(DT,5,2) + '-' + Copy(DT,7,2);
      12: Result:= Copy(DT,1,4) + '-' + Copy(DT,5,2) + '-' + Copy(DT,7,2) + ' ' + Copy(DT,9,2) + ':' + Copy(DT,11,2);
      14: Result:= Copy(DT,1,4) + '-' + Copy(DT,5,2) + '-' + Copy(DT,7,2) + ' ' + Copy(DT,9,2) + ':' + Copy(DT,11,2) + ':' + Copy(DT,13,2);
      else
          Result:= '';
  end;
end;

function GetSampleDateTime(cData:string):TDateTime;
var
  yyyy,mm,dd,hh,nn,ss:string;
  Date,Time:string;
begin
  Result:= now;
  //20080326112638
  //2008060911222315
  if Length(cData) < 14 then exit;

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
  for i:=1 to TGrid.ColCount-1 do begin
      if Abbr = TGrid.Cells[i,0] then begin
          Result:= i;
          exit;
      end;
  end;

end;

function AddRowIndex(var TGrid:TAdvStringGrid):integer;
begin
  Result:= 1;

  if TGrid.Cells[0,1] <> '' then
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
  if UpperCase(Value) = 'TRUE' then
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
var
  OsLocale:integer;
begin
  AppPath:= ExtractFilePath(Application.ExeName);
  OsLocale:= GetSystemDefaultLCID;
  Case OsLocale of
    Ord(KOR): OsLang:= KOR;
    Ord(ENG): OsLang:= ENG;
    else
        OsLang:= KOR;
  end;

  IPAddr:= GetIpAddr;

  LoadIni;

  ClearLog;
end;

procedure TGlobalVar.LoadIni;
var
  TIni: TIniFile;
begin
  tIni:= TIniFile.Create(AppPath+ IniFileName);
  try
      with tIni do begin
          FInstName:= ReadString('INST', 'NAME', '');
          FUserId  := ReadString('USER', 'ID', '');
          MainTop   := ReadInteger('DISP', 'TOP', DEFTOP);
          MainLeft  := ReadInteger('DISP', 'LFT', DEFLFT);
          MainWidth := ReadInteger('DISP', 'WIDTH', DEFWID);
          MainHeigh := ReadInteger('DISP', 'HEIGH', DEFHEI);
          AppTitle  := ReadString('DISP', 'TITLE', '');
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
          WriteInteger('DISP', 'TOP', MainTop);
          WriteInteger('DISP', 'LFT', MainLeft);
          WriteInteger('DISP', 'WIDTH', MainWidth);
          WriteInteger('DISP', 'HEIGH', MainHeigh);
          WriteString('DISP', 'TITLE', AppTitle);
      end;

  finally
      tIni.Free;
  end;

end;

procedure TGlobalVar.SetLogMsg(cValue:string);
begin
  WriteLog(LogFileName, cValue);
end;

procedure TGlobalVar.SetSaveLog(const Value: string);
begin
  WriteLog(RCVFileName, Value);
end;

procedure TGlobalVar.WriteLog(cFileName:string; cStr: string);
var
  F:TextFile;
begin
try
  AssignFile(F,cFileName);

  if FileExists(cFileName) = False then
      exit //Rewrite(F)
  else
      Append(F);

  try
      Writeln(F, #13#10 + FormatDateTime('yyyy-mm-dd hh:nn:ss', now) + #13#10 + cStr);

  finally
      CloseFile(F);
  end;
except
end;
end;

procedure TGlobalVar.ComPortIniLoad;
var
  TIni: TIniFile;
begin
  tIni:= TIniFile.Create(AppPath+ IniFileName);
  try
      with tIni, ComPortSet do begin
          PortNum   := ReadInteger('MSCOMM1', 'Port', 1);
          BaudRate  := ReadString('MSCOMM1', 'BaudRate', '9600');
          Parity    := ReadString('MSCOMM1', 'Parity', 'None');
          DataBit   := ReadString('MSCOMM1', 'DataBits', '8');
          StopBit   := ReadString('MSCOMM1', 'StopBits', '1');
          HandShake := ReadInteger('MSCOMM1', 'FlowControl', 0);
          Dtr       := ReadBool('MSCOMM1', 'DTR', True);
          Rts       := ReadBool('MSCOMM1', 'RTS', True);

          Settings:= BaudRate + ',' + Copy(Parity,1,1) + ',' + Databit + ',' + StopBit;
      end;

  finally
      tIni.Free;
  end;

end;

procedure TGlobalVar.ComPortIniSave;
var
  TIni: TIniFile;
begin
  tIni:= TIniFile.Create(AppPath+ IniFileName);
  try
      with tIni, ComPortSet do begin
          WriteInteger('MSCOMM1', 'Port', PortNum);
          WriteString('MSCOMM1', 'BaudRate', BaudRate);
          WriteString('MSCOMM1', 'Parity', Parity);
          WriteString('MSCOMM1', 'DataBits', DataBit);
          WriteString('MSCOMM1', 'StopBits', StopBit);
          WriteInteger('MSCOMM1', 'FlowControl', HandShake);
          WriteBool('MSCOMM1', 'DTR', Dtr);
          WriteBool('MSCOMM1', 'RTS', Rts);
      end;

  finally
      tIni.Free;
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

destructor TGlobalVar.Destroy;
begin
  SaveIni;
  inherited;
end;

function TGlobalVar.GetIpAddr: string;
var
   HostName  : PChar;
   pHostEnt_ : PHostEnt;
   wVersionRequested : WORD;
   wsaData : TWSADATA;
   szAddr : String;
begin
     Result := '';
     wVersionRequested := MAKEWORD(1, 1);
     HostName := nil;
     if (WSAStartup(wVersionRequested, wsaData) <> 0) then exit;
     try
        HostName := AllocMem(255);
        if (Winsock.gethostname(Hostname, 255) <> 0) then exit;
        pHostEnt_  := Winsock.gethostbyname(HostName);
        if (pHostEnt_ = nil) then exit;
        szAddr := IntToStr(Ord((pHostEnt_^.h_addr_list^)^)) + '.';
        szAddr := szAddr + IntToStr(Ord((pHostEnt_^.h_addr_list^ +1)^)) + '.';
        szAddr := szAddr + IntToStr(Ord((pHostEnt_^.h_addr_list^ +2)^)) + '.';
        szAddr := szAddr + IntToStr(Ord((pHostEnt_^.h_addr_list^ +3)^));
        Result := szAddr;
     finally
            if (HostName <> nil) then FreeMem(HostName);
     end;
end;


procedure TGlobalVar.ClearLog;
var
  i:integer;
  slLog:TStringList;
begin
  //sDate:= FormatDateTime('yyyy-mm-dd', now -10);

  slLog:= TStringList.Create;
  try
      try
          //RCV.Log
          slLog.LoadFromFile(DataFileName);
          for i:= slLog.Count -1 downto 0 do begin
              if StrToDateDef(Copy(slLog.Strings[i],1,10), now) < (Now-10) then begin
                  slLog.Delete(i); slLog.Delete(i+1);  slLog.Delete(i+2);
              end;
          end;
          slLog.SaveToFile(DataFileName);

          //Error.Log
          slLog.LoadFromFile(LogFileName);
          for i:= slLog.Count -1 downto 0 do begin
              if StrToDateDef(Copy(slLog.Strings[i],1,10), now) < (Now-10) then begin
                  slLog.Delete(i); slLog.Delete(i+1); slLog.Delete(i+2);
              end;
          end;
          slLog.SaveToFile(LogFileName);

      except
        on e:exception do begin
            ShowMessage('???????????? ????->'+e.Message);
        end;
      end;

  finally
    slLog.Free;
  end;

end;

end.
