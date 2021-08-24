unit TuxClient;

interface

uses Classes;

const
  TuxedoDLLName = 'C:\cuh_95\WTUXWS32.DLL'; //화순:WTUXWS32.DLL
  //FML32
  {
  TP_BITMAP  = 201326683;
  Pseudo1_32 = 167779262;
  Pseudo2_32 = 167779263;
  Pseudo3_32 = 167779264;
  Pseudo4_32 = 167779265;
  Pseudo5_32 = 167779266;
  Pseudo6_32 = 167779267;
  Pseudo7_32 = 167779268;
  Pseudo8_32 = 167779269;
  Pseudo9_32 = 167779270;
  Pseudo10_32 = 167779271;
  Pseudo11_32 = 167779272;
  Pseudo12_32 = 167779273;
  Pseudo13_32 = 167779274;
  Pseudo14_32 = 167779275;
  Pseudo15_32 = 167779276;
  Pseudo16_32 = 167779277;
  Pseudo17_32 = 167779278;
  Pseudo18_32 = 167779279;
  Pseudo19_32 = 167779280;
  Pseudo20_32 = 167779281;

  //FML16
  Pseudo1	 = 48062;
  Pseudo2	 = 48063;	
  Pseudo3	 = 48064;	
  Pseudo4	 = 48065;	
  Pseudo5	 = 48066;	
  Pseudo6	 = 48067;	
  Pseudo7	 = 48068;	
  Pseudo8	 = 48069;	
  Pseudo9	 = 48070;	
  Pseudo10 = 48071;
  Pseudo11 = 48072;
  Pseudo12 = 48073; }
  
  //MSG_TEXT = 167820962;	  //number: 48802	 type: string

  MAXTIDENT	= 30;		{ max len of a /T identifier }
  //STATLIN  = 167772165;   {/* number: 5	 type: string */     }

  FBufferType   = 'FML';         // Buffer Type : FML
  FBufferType32 = 'FML32';       // Buffer Type : FML32
  FBufferSize   = 65532;         // FML Buffer MAX Size (64k)
  FBufferSize32 = 524256;        // FML32 Buffer MAX Size (512k)

  DefaultErrMsg = '전산실 담당자에게 문의하시고, 사용하시기 바랍니다.';
  TPInitSize    = 512;           // tpInit Buffer Size
  TimeOut       = 10;            // 서비스지연 대기시간을 10초로 고정


type
  FIELDID  = Word;
  FLDLEN   = Word;
  FLDOCC   = Integer;
  pFIELDID = ^FIELDID;
  pFLDLEN  = ^FLDLEN;
  pFLDOCC  = ^FLDOCC;

  FIELDID32  = LongInt;
  FLDLEN32   = LongInt;
  FLDOCC32   = LongInt;
  pFIELDID32 = ^FIELDID32;
  pFLDLEN32  = ^FLDLEN32;
  pFLDOCC32  = ^FLDOCC32;


  tpinfo_t = record
     usrname: array[0..MAXTIDENT + 1] of char;	{ client user name }
     cltname: array[0..MAXTIDENT + 1] of char;	{ application client name }
     passwd:  array[0..MAXTIDENT + 1] of char;	{ application password }
     grpname: array[0..MAXTIDENT + 1] of char;	{ client group name }
     flags:   LongInt;		                { initialization flags }
     datalen: LongInt;		                { length of app specific data }
     data:    LongInt;		                { placeholder for app data }
   end;
   pTPINIT = ^tpinfo_t;

var
  transf : Pointer;          // FML Buffer Pointer.

//FML
Function Finit(a:Pointer; b:FLDLEN):Integer; stdcall; far;
Function Finit32(a:Pointer; b:FLDLEN32):Integer; stdcall; far;
Function CFchg(a:Pointer; b:FIELDID; c:FLDOCC; d:Pointer; e:FLDLEN; f:Integer):Integer; stdcall; far;
Function Fchg(a:Pointer; b:FIELDID; c:FLDOCC; d:Pointer; e:FLDLEN):Integer; stdcall; far;
Function Fchg32(a:Pointer; b:FIELDID32; c:FLDOCC32; d:Pointer; e:FLDLEN32):Integer; stdcall; far;
Function CFchg32(a:Pointer; b:FIELDID32; c:FLDOCC32; d:Pointer; e:FLDLEN32; f:Integer):Integer; stdcall; far;
Function Fvals(a:Pointer; b:FIELDID; c:FLDOCC):PChar; stdcall; far;
Function getFerror32:Integer; stdcall; far;
Function Fstrerror32(a:Integer):PChar; stdcall; far;
Function Fldid(a:PChar):FIELDID; stdcall; far;
function Fldid32(name:PChar):FIELDID32; stdcall; far;
Function Foccur(a:Pointer; b:FIELDID):FLDOCC; stdcall; far;
function Fvals32(fbfr:Longint; fieldid:Longword; oc:Longint):PChar; stdcall; far;


//Put할때의 총 필드 바이트수 리턴. f:1회 put할때 사용되는 field수, v:1회 put할때 사용되는 총 byte 수
function Fneeded32(f:Longint; v:Longword):Longint; stdcall; far;
//ATMI

Function gettperrno:Integer; stdcall; far;
Function gettpurcode:LongInt; stdcall; far;

Function tuxreadenv(a:PChar; b:PChar):Integer; stdcall; far;
Function tpinit(a:pTPINIT):Integer; stdcall; far;
Function tpalloc(a:PChar; b:PChar; c:LongInt):Pointer; stdcall; far;
Function tpbegin(a:LongInt; b:LongInt):Integer; stdcall; far;
Function tpcall(a:PChar; b:Pointer; c:LongInt; d:Pointer; e:pLongInt; f:LongInt):Integer; stdcall; far;
Function tpcommit(a:LongInt):Integer; stdcall; far;
Function tpstrerror(a:Integer):PChar; stdcall; far;
Function tpabort(a:LongInt):Integer; stdcall; far;
Procedure tpfree(a:Pointer); stdcall; far;
Function tpterm:Integer; stdcall; far;

// 9.3 Image Data Sending procedure. FML
procedure SetFldImage(OccurNo:Integer;FldNameStr:String; ImgStream:TMemoryStream);

implementation

//FML
Function Finit(a:Pointer; b:FLDLEN):Integer; stdcall; external TuxedoDLLName;
Function Finit32(a:Pointer; b:FLDLEN32):Integer; stdcall; external TuxedoDLLName;
Function Fchg(a:Pointer; b:FIELDID; c:FLDOCC; d:Pointer; e:FLDLEN):Integer; stdcall; external TuxedoDLLName;
Function CFchg(a:Pointer; b:FIELDID; c:FLDOCC; d:Pointer; e:FLDLEN; f:Integer):Integer; stdcall; external TuxedoDLLName;
Function Fchg32(a:Pointer; b:FIELDID32; c:FLDOCC32; d:Pointer; e:FLDLEN32):Integer; stdcall; external TuxedoDLLName;
Function CFchg32(a:Pointer; b:FIELDID32; c:FLDOCC32; d:Pointer; e:FLDLEN32; f:Integer):Integer; stdcall; external TuxedoDLLName;
Function Fvals(a:Pointer; b:FIELDID; c:FLDOCC):PChar; stdcall; external TuxedoDLLName;
Function Fstrerror32(a:Integer):PChar; stdcall; external TuxedoDLLName;
Function getFerror32:Integer; stdcall; external TuxedoDLLName;
Function Fldid(a:PChar):FIELDID; stdcall; external TuxedoDLLName;
function Fldid32(name:PChar):FIELDID32; stdcall; external TuxedoDLLName;
Function Foccur(a:Pointer; b:FIELDID):FLDOCC; stdcall; external TuxedoDLLName;
function Fvals32(fbfr:Longint; fieldid:Longword; oc:Longint):PChar; stdcall; external TuxedoDLLName;
//Put할때의 총 필드 바이트수 리턴. f:1회 put할때 사용되는 field수, v:1회 put할때 사용되는 총 byte 수
function Fneeded32(f:Longint; v:Longword):Longint; stdcall; external TuxedoDLLName;
//ATMI
Function gettperrno:Integer; stdcall; external TuxedoDLLName;
Function gettpurcode:LongInt; stdcall; external TuxedoDLLName;

Function tuxreadenv(a:PChar; b:PChar):Integer; stdcall; external TuxedoDLLName;
Function tpinit(a:pTPINIT):Integer; stdcall; external TuxedoDLLName;
Function tpalloc(a:PChar; b:PChar; c:LongInt):Pointer; stdcall; external TuxedoDLLName;
Function tpbegin(a:LongInt; b:LongInt):Integer; stdcall; external TuxedoDLLName;
Function tpcall(a:PChar; b:Pointer; c:LongInt; d:Pointer; e:pLongInt; f:LongInt):Integer; stdcall; external TuxedoDLLName;
Function tpcommit(a:LongInt):Integer; stdcall; external TuxedoDLLName;
Function tpstrerror(a:Integer):PChar; stdcall; external TuxedoDLLName;
Function tpabort(a:LongInt):Integer; stdcall; external TuxedoDLLName;
Procedure tpfree(a:Pointer); stdcall; external TuxedoDLLName;
Function tpterm:Integer; stdcall; external TuxedoDLLName;

// 9.3 Image Data Sending procedure. FML
procedure SetFldImage(OccurNo:Integer;FldNameStr:String; ImgStream:TMemoryStream);
begin
     if (transf <> nil) then
        FChg(Transf,Fldid(PChar(FldNameStr)),OccurNo,ImgStream.Memory,
             ImgStream.Size);
end;





end.
