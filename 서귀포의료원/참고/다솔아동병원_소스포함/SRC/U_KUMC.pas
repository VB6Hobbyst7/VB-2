unit U_KUMC;

interface
uses
  SysUtils, Windows, Classes, Forms, Variants, dialogs;



function TuxedoInit_V2(in_usrname,in_cltname, in_svrid: pChar): Integer; StdCall; external 'C:\DEV\DLL\P_SLDLL.dll';
    //-	DR 환경 구성에 따른 호출 인자 추가 (사용자 로그인 창에 추가 요청)
    //	(정상운영시 ‘01’, DR 상황 발생시 ‘DR’
    //- in_usrname, in_cltname null로 보내도 됨
    // 예) if TuxedoInit_V2('','','01')  = 1 이면 TMAX 연결 OK


function UserChk(in_userid, in_pass, in_locate : string ; var out_usernm: variant): Integer; StdCall; external 'C:\DEV\DLL\P_SLDLL.dll';
   //-  in_userid <- 사번, in_pass <- 패스워드, in_locate <- 병원구분(AA,GR,AS로 입력: AA(안암),GR(구로),AS(안산)
   //-  Return 값이 1이면, out_usernm 에 사용자 이름
   //   (5분마다 타이머로 UserChk 함수 호출로 세션 끊김 방지 처리)

function AutoAccept_poct(in_spcid, in_userid: String; var out_msg :string):integer; StdCall external 'C:\DEV\DLL\P_SLDLL.dll';
   //- 검체를 채취(B) 상태에서 결과입력상태(E)로 변경
   //-  AutoAccept_POCT('102020202',’88888’, out_msg);
   //                   ----------
   //                    검체번호
   //    성공 시 --> return value  1 & out_msg = 'OK'
   //    실패 시 --> return value -1 & out_msg 는 에러 관련 메시지

function ExaminfoList(in_Flag, in_spcid, in_execdate: string ; var out_order: variant) : Integer; StdCall;  external 'C:\DEV\DLL\P_SLDLL.dll';
    {-  ExaminfoList('', '1231231231', '20080908', out_msg);

       성공적으로 조회 시 out_msg 다음과 같은 형식으로 보내줌 , return value 는 검사 건수
       구분자를 '|' 로 구분함.
       examinfo.sPatno   [ii] + '|' +    --> 환자 등록번호
       examinfo.sPatname [ii] + '|' +    --> 환자 명
       examinfo.sSex     [ii] + '|' +    --> 성별 (F,M)
       examinfo.sAge     [ii] + '|' +    --> 나이
       examinfo.sOrdseqno[ii] + '|' +    --> 처방순번
       examinfo.sWorkno  [ii] + '|' +    --> 작업번호
       examinfo.sOrddate [ii] + '|' +    --> 처방일자(yyyy-mm-dd')
       examinfo.sExamcode[ii] + '|' +    --> 검사코드
       examinfo.sAcptdate[ii] + '|' +    --> 접수일자
       examinfo.sAreano  [ii] + '|' +    --> 접수번호
       examinfo.sMeddept [ii] + '|' +    --> 진료과
       examinfo.sSpccode [ii] + '|' +    --> 검체코드
       examinfo.sSpcname [ii];           --> 검체명

       예)
       00001234|홍길동|F|62|2001|1|2008-09-08|BG2200|2008-09-08|1|OG|101|Whole Blood
    }

function ResultList(out_jobsect, out_Userid:string; out_Result : variant ; out_eqipcd: string='';out_autoryn: string='N') : Integer; StdCall external 'C:\DEV\DLL\P_SLDLL.dll';
{
       ResultList('3', '88888', out_Result, 'EQIP_1', 'N');
       '2'     --> 보고처리 ( 결과만 입력인 경우에는 '3' 으로 처리)
       '99999' --> login 사번
       'EQIP_1' --> 검사 실시한 장비코드
       out_result 처리 방법
       검사결과를 '|' 로 구분하여 Upload 처리
       검체번호|등록번호|처방일자|처방순번|검사코드|검사결과|검사특기시항|delta|Panic|정상치flag|장비flag|
       **       **       **       **       **        **
       순으로 보내줌 ( ** 는 필수 입력사항임)
       
       예)
       8080004709|80000412|2008-07-28|2001|BG23201A01|2|||||
       8080004709|80000412|2008-07-28|2001|BG23201A02|2|||||
       8080004709|80000412|2008-07-28|2001|BG23201A03|0|||||
       
** PC에 에러 및 로그를 파일로 저장하여 주시기 바람.
}

/////
function DownLoadOrder_KUMC_Result_Test(TOBJ:TObject): boolean;
function Con_KUMC(STA:string):integer;

implementation

uses
  GlobalVar, U_IfClass, SetDataBase, StringLib;

function Con_KUMC(STA:string):integer;
begin
  Result:= TuxedoInit_V2(PChar(''), PChar(''), PChar(STA));
end;

function DownLoadOrder_KUMC_Result_Test(TOBJ:TObject): boolean;
var
  TMaster:TIfMaster absolute TOBJ;
  vOrder:variant;
  OrdList,sCode : String;
  R, i, j: integer;
  sHttp, SndStr:string;
begin
  Result:= False;
  if SvrConnection = False then
      exit;

  R:= ExaminfoList('', TMaster.FBarCode, FormatDateTime('yyyymmdd', now),vOrder);

  if R < 1 then exit;

  Result:= True;

  TMaster.vOrdList:= vOrder;

  exit;

  {
  for i:=0 to R-1 do begin
      OrdList:= vOrder[i];
      TMaster.FPID:= TokenStr( OrdList, '|', 1);           //--> 환자 등록번호 (PK)
      TMaster.FPNM:= TokenStr( OrdList, '|', 2);           //--> 환자 명
      TMaster.FSex:= TokenStr( OrdList, '|', 3);           //--> 성별 (F,M)
      TMaster.FAge:= TokenStr( OrdList, '|', 4);           //--> 나이
      TMaster.FOrdSeq:= TokenStr( OrdList, '|', 5);        //--> 처방순번      (PK)
      TMaster.FWorkNo:= TokenStr( OrdList, '|', 6);        //--> 작업번호
      TMaster.FOrdDate:= TokenStr( OrdList, '|', 7);       //--> 처방일자(yyyy-mm-dd')  (PK)
      TMaster.FExamCode:= TokenStr( OrdList, '|', 8);      //--> 검사코드      (PK)
      TMaster.FAcptDt:= TokenStr( OrdList, '|', 9);        //--> 접수일자
      //examinfo.sAreano:= TokenStr( OrdList, '|', 10);    //--> 접수번호
      //examinfo.sMeddept:= TokenStr( OrdList, '|', 11);   //--> 진료과
      //examinfo.sSpccode:= TokenStr( OrdList, '|', 12);   //--> 검체코드
      //examinfo.sSpcname:= TokenStr( OrdList, '|', 13);   //--> 검체명
  end;}

end;

end.
