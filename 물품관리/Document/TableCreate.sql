
CREATE TABLE buyL
(
	buydt                varchar(10)  NOT NULL ,
	buyseq               smallint  NOT NULL ,
	stkcd                integer  NULL ,
	buyqty               numeric(8,1)  NULL ,
	qtyrate              numeric(5)  NULL ,
	ioqty                numeric(8,1)  NULL ,
	amt                  numeric(11,1)  NULL ,
	sumamt               numeric(12)  NULL ,
	maxdt                varchar(10)  NULL ,
	makeno               varchar(30)  NULL ,
	custcd               numeric(5)  NULL ,
	usercd               varchar(10)  NULL ,
	ordym                varchar(6)  NULL ,
	ordno                smallint  NULL ,
	ordseq               smallint  NULL ,
	wrtdt                varchar(20)  NULL ,
	moddt                varchar(20)  NULL 
)
go


ALTER TABLE buyL
	ADD CONSTRAINT XPK�����԰� PRIMARY KEY  CLUSTERED (buydt ASC,buyseq ASC)
go


CREATE NONCLUSTERED INDEX XIF1�����԰� ON buyL
(
	stkcd                 ASC
)
go


CREATE NONCLUSTERED INDEX XIF3�����԰� ON buyL
(
	usercd                ASC
)
go


CREATE NONCLUSTERED INDEX XIF4�����԰� ON buyL
(
	ordym                 ASC,
	ordno                 ASC,
	ordseq                ASC
)
go


CREATE NONCLUSTERED INDEX XIF2�����԰� ON buyL
(
	custcd                ASC
)
go


CREATE TABLE dayMENU
(
	testdt               varchar(10)  NOT NULL ,
	testseq              smallint  NOT NULL ,
	testcd               varchar(7)  NULL ,
	testcnt              numeric(3)  NULL ,
	reason               varchar(30)  NULL ,
	usercd               varchar(10)  NULL ,
	endfg                numeric(1)  NULL ,
	wrtdt                varchar(20)  NULL ,
	moddt                varchar(20)  NULL 
)
go


ALTER TABLE dayMENU
	ADD CONSTRAINT XPK�����˻���� PRIMARY KEY  CLUSTERED (testdt ASC,testseq ASC)
go


CREATE NONCLUSTERED INDEX XIF1�����˻���� ON dayMENU
(
	testcd                ASC
)
go


CREATE NONCLUSTERED INDEX XIF2�����˻���� ON dayMENU
(
	usercd                ASC
)
go


CREATE TABLE dayTEST
(
	enddt                varchar(10)  NOT NULL ,
	testcd               varchar(7)  NOT NULL ,
	testcnt              numeric(8)  NULL ,
	manucnt              numeric(8)  NULL 
)
go


ALTER TABLE dayTEST
	ADD CONSTRAINT XPK�˻縶�� PRIMARY KEY  CLUSTERED (enddt ASC,testcd ASC)
go


CREATE NONCLUSTERED INDEX XIF1�˻縶�� ON dayTEST
(
	testcd                ASC
)
go


CREATE TABLE machSTK
(
	machcd               varchar(3)  NOT NULL ,
	opercd               numeric(2)  NOT NULL ,
	stkcd                integer  NOT NULL ,
	qty                  numeric(8,1)  NULL 
)
go


ALTER TABLE machSTK
	ADD CONSTRAINT XPK��񺰽þ೻�� PRIMARY KEY  CLUSTERED (machcd ASC,opercd ASC,stkcd ASC)
go


CREATE NONCLUSTERED INDEX XIF1��񺰽þ೻�� ON machSTK
(
	machcd                ASC
)
go


CREATE NONCLUSTERED INDEX XIF2��񺰽þ೻�� ON machSTK
(
	opercd                ASC
)
go


CREATE NONCLUSTERED INDEX XIF3��񺰽þ೻�� ON machSTK
(
	stkcd                 ASC
)
go


CREATE TABLE mstCUST
(
	custcd               numeric(5)  NOT NULL ,
	custnm               varchar(30)  NULL ,
	custmng              varchar(20)  NULL ,
	custid               varchar(15)  NULL ,
	custtype             varchar(30)  NULL ,
	custitem             varchar(30)  NULL ,
	addr1                varchar(50)  NULL ,
	addr2                varchar(50)  NULL ,
	postno               varchar(7)  NULL ,
	telno                varchar(15)  NULL ,
	faxno                varchar(15)  NULL ,
	custman              varchar(20)  NULL ,
	hpno                 varchar(15)  NULL ,
	banknm               varchar(30)  NULL ,
	bankno               varchar(30)  NULL ,
	remark               varchar(200)  NULL ,
	delfg                numeric(1)  NULL ,
	wrtdt                varchar(20)  NULL ,
	moddt                varchar(20)  NULL 
)
go


ALTER TABLE mstCUST
	ADD CONSTRAINT XPK����ó���� PRIMARY KEY  CLUSTERED (custcd ASC)
go


CREATE TABLE mstDUTY
(
	dutycd               varchar(2)  NOT NULL ,
	dutynm               varchar(30)  NULL 
)
go


ALTER TABLE mstDUTY
	ADD CONSTRAINT XPK�μ����� PRIMARY KEY  CLUSTERED (dutycd ASC)
go


CREATE TABLE mstMACH
(
	machcd               varchar(3)  NOT NULL ,
	machnm               varchar(30)  NULL ,
	delfg                numeric(1)  NULL ,
	dutycd               varchar(2)  NULL 
)
go


ALTER TABLE mstMACH
	ADD CONSTRAINT XPK������ PRIMARY KEY  CLUSTERED (machcd ASC)
go


CREATE NONCLUSTERED INDEX XIF1������ ON mstMACH
(
	dutycd                ASC
)
go


CREATE TABLE mstOPER
(
	opercd               numeric(2)  NOT NULL ,
	opernm               varchar(30)  NULL ,
	operfg               numeric(1)  NULL 
)
go


ALTER TABLE mstOPER
	ADD CONSTRAINT XPK������� PRIMARY KEY  CLUSTERED (opercd ASC)
go


CREATE TABLE mstSTK
(
	stkcd                integer  NOT NULL ,
	stknm                varchar(50)  NULL ,
	stkspec              varchar(50)  NULL ,
	maker                varchar(50)  NULL ,
	buyunit              varchar(5)  NULL ,
	iounit               varchar(5)  NULL ,
	buyioqty             numeric(5)  NULL ,
	stdamt               numeric(11,1)  NULL ,
	buyamt               numeric(11,1)  NULL ,
	buyday               numeric(3)  NULL ,
	minbuyqty            numeric(8,1)  NULL ,
	barcode              varchar(20)  NULL ,
	rmdfg                numeric(1)  NULL ,
	safeqty              numeric(8,1)  NULL ,
	buytype              numeric(1)  NULL ,
	kindcd               varchar(2)  NULL ,
	custcd               numeric(5)  NULL ,
	delfg                numeric(1)  NULL ,
	wrtdt                varchar(20)  NULL ,
	moddt                varchar(20)  NULL 
)
go


ALTER TABLE mstSTK
	ADD CONSTRAINT XPK��ǰ���� PRIMARY KEY  CLUSTERED (stkcd ASC)
go


CREATE NONCLUSTERED INDEX XIF1��ǰ���� ON mstSTK
(
	kindcd                ASC
)
go


CREATE NONCLUSTERED INDEX XIE1��ǰ���� ON mstSTK
(
	stknm                 ASC
)
go


CREATE NONCLUSTERED INDEX XIE2��ǰ���� ON mstSTK
(
	kindcd                ASC,
	stkcd                 ASC
)
go


CREATE NONCLUSTERED INDEX XIF2��ǰ���� ON mstSTK
(
	custcd                ASC
)
go


CREATE TABLE mstSTKG
(
	kindcd               varchar(2)  NOT NULL ,
	kindnm               varchar(30)  NULL ,
	reagentfg            numeric(1)  NULL 
)
go


ALTER TABLE mstSTKG
	ADD CONSTRAINT XPK��ǰ�з� PRIMARY KEY  CLUSTERED (kindcd ASC)
go


CREATE TABLE mstTEST
(
	testcd               varchar(7)  NOT NULL ,
	testnm               varchar(30)  NULL 
)
go


ALTER TABLE mstTEST
	ADD CONSTRAINT XPK�˻��׸� PRIMARY KEY  CLUSTERED (testcd ASC)
go


CREATE TABLE mstUSER
(
	usercd               varchar(10)  NOT NULL ,
	usernm               varchar(20)  NULL ,
	pswd                 varchar(15)  NULL ,
	levelfg              numeric(1)  NULL ,
	delfg                numeric(1)  NULL ,
	dutycd               varchar(2)  NULL ,
	wrtdt                varchar(20)  NULL ,
	moddt                varchar(20)  NULL 
)
go


ALTER TABLE mstUSER
	ADD CONSTRAINT XPK����ڱ��� PRIMARY KEY  CLUSTERED (usercd ASC)
go


CREATE NONCLUSTERED INDEX XIF1����ڱ��� ON mstUSER
(
	dutycd                ASC
)
go


CREATE TABLE operL
(
	machcd               varchar(3)  NOT NULL ,
	operdt               varchar(10)  NOT NULL ,
	operseq              smallint  NOT NULL ,
	opercd               numeric(2)  NULL ,
	opercnt              smallint  NULL ,
	endfg                numeric(1)  NULL ,
	reason               varchar(30)  NULL ,
	usercd               varchar(10)  NULL ,
	wrtdt                varchar(20)  NULL ,
	moddt                varchar(20)  NULL 
)
go


ALTER TABLE operL
	ADD CONSTRAINT XPK������� PRIMARY KEY  CLUSTERED (machcd ASC,operdt ASC,operseq ASC)
go


CREATE NONCLUSTERED INDEX XIF1������� ON operL
(
	machcd                ASC
)
go


CREATE NONCLUSTERED INDEX XIF2������� ON operL
(
	opercd                ASC
)
go


CREATE NONCLUSTERED INDEX XIF3������� ON operL
(
	usercd                ASC
)
go


CREATE TABLE ordH
(
	ordym                varchar(6)  NOT NULL ,
	ordno                smallint  NOT NULL ,
	orddt                varchar(10)  NULL ,
	ordtype              numeric(1)  NULL ,
	ordamt               numeric(12)  NULL ,
	remark               varchar(30)  NULL ,
	custcd               numeric(5)  NULL ,
	usercd               varchar(10)  NULL ,
	wrtdt                varchar(20)  NULL ,
	moddt                varchar(20)  NULL 
)
go


ALTER TABLE ordH
	ADD CONSTRAINT XPK���ּ� PRIMARY KEY  CLUSTERED (ordym ASC,ordno ASC)
go


CREATE NONCLUSTERED INDEX XIF2���ּ� ON ordH
(
	usercd                ASC
)
go


CREATE NONCLUSTERED INDEX XIF1���ּ� ON ordH
(
	custcd                ASC
)
go


CREATE TABLE ordL
(
	ordym                varchar(6)  NOT NULL ,
	ordno                smallint  NOT NULL ,
	ordseq               smallint  NOT NULL ,
	stkcd                integer  NULL ,
	qty                  numeric(8,1)  NULL ,
	amt                  numeric(11,1)  NULL ,
	sumamt               numeric(12)  NULL ,
	lastdt               varchar(10)  NULL ,
	inqty                numeric(8,1)  NULL ,
	lastindt             varchar(10)  NULL ,
	remark               varchar(30)  NULL 
)
go


ALTER TABLE ordL
	ADD CONSTRAINT XPK���ֳ��� PRIMARY KEY  CLUSTERED (ordym ASC,ordno ASC,ordseq ASC)
go


CREATE NONCLUSTERED INDEX XIF1���ֳ��� ON ordL
(
	ordym                 ASC,
	ordno                 ASC
)
go


CREATE NONCLUSTERED INDEX XIF2���ֳ��� ON ordL
(
	stkcd                 ASC
)
go


CREATE TABLE outL
(
	outdt                varchar(10)  NOT NULL ,
	outfg                numeric(1)  NOT NULL ,
	outseq               smallint  NOT NULL ,
	stkcd                integer  NULL ,
	qty                  numeric(8,1)  NULL ,
	reason               varchar(30)  NULL ,
	dutycd               varchar(2)  NULL ,
	usercd               varchar(10)  NULL ,
	wrtdt                varchar(20)  NULL ,
	moddt                varchar(20)  NULL 
)
go


ALTER TABLE outL
	ADD CONSTRAINT XPK��� PRIMARY KEY  CLUSTERED (outdt ASC,outfg ASC,outseq ASC)
go


CREATE NONCLUSTERED INDEX XIF1��� ON outL
(
	stkcd                 ASC
)
go


CREATE NONCLUSTERED INDEX XIF2��� ON outL
(
	usercd                ASC
)
go


CREATE NONCLUSTERED INDEX XIE1��� ON outL
(
	outdt                 ASC,
	stkcd                 ASC
)
go

CREATE NONCLUSTERED INDEX XIF3��� ON outL
(
	dutycd                ASC
)
go


CREATE TABLE reqL
(
	dutycd               varchar(2)  NOT NULL ,
	reqdt                varchar(10)  NOT NULL ,
	reqseq               smallint  NOT NULL ,
	stkcd                integer  NULL ,
	qty                  numeric(8,1)  NULL ,
	stat                 numeric(1)  NULL ,
	lastdt               varchar(10)  NULL ,
	remark               varchar(30)  NULL ,
	usercd               varchar(10)  NULL ,
	ordym                varchar(6)  NULL ,
	ordno                smallint  NULL ,
	ordseq               smallint  NULL ,
	wrtdt                varchar(20)  NULL ,
	moddt                varchar(20)  NULL 
)
go


ALTER TABLE reqL
	ADD CONSTRAINT XPK���ſ�û�� PRIMARY KEY  CLUSTERED (dutycd ASC,reqdt ASC,reqseq ASC)
go


CREATE NONCLUSTERED INDEX XIF1���ſ�û�� ON reqL
(
	dutycd                ASC
)
go


CREATE NONCLUSTERED INDEX XIF2���ſ�û�� ON reqL
(
	stkcd                 ASC
)
go


CREATE NONCLUSTERED INDEX XIF3���ſ�û�� ON reqL
(
	ordym                 ASC,
	ordno                 ASC,
	ordseq                ASC
)
go


CREATE NONCLUSTERED INDEX XIF4���ſ�û�� ON reqL
(
	usercd                ASC
)
go


CREATE TABLE stkRMD
(
	stkcd                integer  NOT NULL ,
	rmdym                varchar(7)  NOT NULL ,
	prevqty              numeric(8,1)  NULL ,
	buyqty               numeric(8,1)  NULL ,
	inqty                numeric(8,1)  NULL ,
	outqty               numeric(8,1)  NULL 
)
go


ALTER TABLE stkRMD
	ADD CONSTRAINT XPK��ǰ��� PRIMARY KEY  CLUSTERED (stkcd ASC,rmdym ASC)
go


CREATE NONCLUSTERED INDEX XIF1��ǰ��� ON stkRMD
(
	stkcd                 ASC
)
go


CREATE TABLE testSTK
(
	testcd               varchar(7)  NOT NULL ,
	stkcd                integer  NOT NULL ,
	qty                  numeric(8,1)  NULL 
)
go


ALTER TABLE testSTK
	ADD CONSTRAINT XPK�˻纰�þ���� PRIMARY KEY  CLUSTERED (testcd ASC,stkcd ASC)
go


CREATE NONCLUSTERED INDEX XIF1�˻纰�þ���� ON testSTK
(
	testcd                ASC
)
go


CREATE NONCLUSTERED INDEX XIF2�˻纰�þ���� ON testSTK
(
	stkcd                 ASC
)
go

