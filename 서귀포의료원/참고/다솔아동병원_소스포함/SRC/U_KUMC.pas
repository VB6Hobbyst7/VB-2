unit U_KUMC;

interface
uses
  SysUtils, Windows, Classes, Forms, Variants, dialogs;



function TuxedoInit_V2(in_usrname,in_cltname, in_svrid: pChar): Integer; StdCall; external 'C:\DEV\DLL\P_SLDLL.dll';
    //-	DR ȯ�� ������ ���� ȣ�� ���� �߰� (����� �α��� â�� �߰� ��û)
    //	(������ ��01��, DR ��Ȳ �߻��� ��DR��
    //- in_usrname, in_cltname null�� ������ ��
    // ��) if TuxedoInit_V2('','','01')  = 1 �̸� TMAX ���� OK


function UserChk(in_userid, in_pass, in_locate : string ; var out_usernm: variant): Integer; StdCall; external 'C:\DEV\DLL\P_SLDLL.dll';
   //-  in_userid <- ���, in_pass <- �н�����, in_locate <- ��������(AA,GR,AS�� �Է�: AA(�Ⱦ�),GR(����),AS(�Ȼ�)
   //-  Return ���� 1�̸�, out_usernm �� ����� �̸�
   //   (5�и��� Ÿ�̸ӷ� UserChk �Լ� ȣ��� ���� ���� ���� ó��)

function AutoAccept_poct(in_spcid, in_userid: String; var out_msg :string):integer; StdCall external 'C:\DEV\DLL\P_SLDLL.dll';
   //- ��ü�� ä��(B) ���¿��� ����Է»���(E)�� ����
   //-  AutoAccept_POCT('102020202',��88888��, out_msg);
   //                   ----------
   //                    ��ü��ȣ
   //    ���� �� --> return value  1 & out_msg = 'OK'
   //    ���� �� --> return value -1 & out_msg �� ���� ���� �޽���

function ExaminfoList(in_Flag, in_spcid, in_execdate: string ; var out_order: variant) : Integer; StdCall;  external 'C:\DEV\DLL\P_SLDLL.dll';
    {-  ExaminfoList('', '1231231231', '20080908', out_msg);

       ���������� ��ȸ �� out_msg ������ ���� �������� ������ , return value �� �˻� �Ǽ�
       �����ڸ� '|' �� ������.
       examinfo.sPatno   [ii] + '|' +    --> ȯ�� ��Ϲ�ȣ
       examinfo.sPatname [ii] + '|' +    --> ȯ�� ��
       examinfo.sSex     [ii] + '|' +    --> ���� (F,M)
       examinfo.sAge     [ii] + '|' +    --> ����
       examinfo.sOrdseqno[ii] + '|' +    --> ó�����
       examinfo.sWorkno  [ii] + '|' +    --> �۾���ȣ
       examinfo.sOrddate [ii] + '|' +    --> ó������(yyyy-mm-dd')
       examinfo.sExamcode[ii] + '|' +    --> �˻��ڵ�
       examinfo.sAcptdate[ii] + '|' +    --> ��������
       examinfo.sAreano  [ii] + '|' +    --> ������ȣ
       examinfo.sMeddept [ii] + '|' +    --> �����
       examinfo.sSpccode [ii] + '|' +    --> ��ü�ڵ�
       examinfo.sSpcname [ii];           --> ��ü��

       ��)
       00001234|ȫ�浿|F|62|2001|1|2008-09-08|BG2200|2008-09-08|1|OG|101|Whole Blood
    }

function ResultList(out_jobsect, out_Userid:string; out_Result : variant ; out_eqipcd: string='';out_autoryn: string='N') : Integer; StdCall external 'C:\DEV\DLL\P_SLDLL.dll';
{
       ResultList('3', '88888', out_Result, 'EQIP_1', 'N');
       '2'     --> ����ó�� ( ����� �Է��� ��쿡�� '3' ���� ó��)
       '99999' --> login ���
       'EQIP_1' --> �˻� �ǽ��� ����ڵ�
       out_result ó�� ���
       �˻����� '|' �� �����Ͽ� Upload ó��
       ��ü��ȣ|��Ϲ�ȣ|ó������|ó�����|�˻��ڵ�|�˻���|�˻�Ư�����|delta|Panic|����ġflag|���flag|
       **       **       **       **       **        **
       ������ ������ ( ** �� �ʼ� �Է»�����)
       
       ��)
       8080004709|80000412|2008-07-28|2001|BG23201A01|2|||||
       8080004709|80000412|2008-07-28|2001|BG23201A02|2|||||
       8080004709|80000412|2008-07-28|2001|BG23201A03|0|||||
       
** PC�� ���� �� �α׸� ���Ϸ� �����Ͽ� �ֽñ� �ٶ�.
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
      TMaster.FPID:= TokenStr( OrdList, '|', 1);           //--> ȯ�� ��Ϲ�ȣ (PK)
      TMaster.FPNM:= TokenStr( OrdList, '|', 2);           //--> ȯ�� ��
      TMaster.FSex:= TokenStr( OrdList, '|', 3);           //--> ���� (F,M)
      TMaster.FAge:= TokenStr( OrdList, '|', 4);           //--> ����
      TMaster.FOrdSeq:= TokenStr( OrdList, '|', 5);        //--> ó�����      (PK)
      TMaster.FWorkNo:= TokenStr( OrdList, '|', 6);        //--> �۾���ȣ
      TMaster.FOrdDate:= TokenStr( OrdList, '|', 7);       //--> ó������(yyyy-mm-dd')  (PK)
      TMaster.FExamCode:= TokenStr( OrdList, '|', 8);      //--> �˻��ڵ�      (PK)
      TMaster.FAcptDt:= TokenStr( OrdList, '|', 9);        //--> ��������
      //examinfo.sAreano:= TokenStr( OrdList, '|', 10);    //--> ������ȣ
      //examinfo.sMeddept:= TokenStr( OrdList, '|', 11);   //--> �����
      //examinfo.sSpccode:= TokenStr( OrdList, '|', 12);   //--> ��ü�ڵ�
      //examinfo.sSpcname:= TokenStr( OrdList, '|', 13);   //--> ��ü��
  end;}

end;

end.
