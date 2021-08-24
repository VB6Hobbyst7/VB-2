
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
	ADD CONSTRAINT XPK구매입고 PRIMARY KEY  CLUSTERED (buydt ASC,buyseq ASC)
go


CREATE NONCLUSTERED INDEX XIF1구매입고 ON buyL
(
	stkcd                 ASC
)
go


CREATE NONCLUSTERED INDEX XIF3구매입고 ON buyL
(
	usercd                ASC
)
go


CREATE NONCLUSTERED INDEX XIF4구매입고 ON buyL
(
	ordym                 ASC,
	ordno                 ASC,
	ordseq                ASC
)
go


CREATE NONCLUSTERED INDEX XIF2구매입고 ON buyL
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
	ADD CONSTRAINT XPK수동검사출고 PRIMARY KEY  CLUSTERED (testdt ASC,testseq ASC)
go


CREATE NONCLUSTERED INDEX XIF1수동검사출고 ON dayMENU
(
	testcd                ASC
)
go


CREATE NONCLUSTERED INDEX XIF2수동검사출고 ON dayMENU
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
	ADD CONSTRAINT XPK검사마감 PRIMARY KEY  CLUSTERED (enddt ASC,testcd ASC)
go


CREATE NONCLUSTERED INDEX XIF1검사마감 ON dayTEST
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
	ADD CONSTRAINT XPK장비별시약내역 PRIMARY KEY  CLUSTERED (machcd ASC,opercd ASC,stkcd ASC)
go


CREATE NONCLUSTERED INDEX XIF1장비별시약내역 ON machSTK
(
	machcd                ASC
)
go


CREATE NONCLUSTERED INDEX XIF2장비별시약내역 ON machSTK
(
	opercd                ASC
)
go


CREATE NONCLUSTERED INDEX XIF3장비별시약내역 ON machSTK
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
	ADD CONSTRAINT XPK구매처기초 PRIMARY KEY  CLUSTERED (custcd ASC)
go


CREATE TABLE mstDUTY
(
	dutycd               varchar(2)  NOT NULL ,
	dutynm               varchar(30)  NULL 
)
go


ALTER TABLE mstDUTY
	ADD CONSTRAINT XPK부서기초 PRIMARY KEY  CLUSTERED (dutycd ASC)
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
	ADD CONSTRAINT XPK장비기초 PRIMARY KEY  CLUSTERED (machcd ASC)
go


CREATE NONCLUSTERED INDEX XIF1장비기초 ON mstMACH
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
	ADD CONSTRAINT XPK장비운영기초 PRIMARY KEY  CLUSTERED (opercd ASC)
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
	ADD CONSTRAINT XPK물품기초 PRIMARY KEY  CLUSTERED (stkcd ASC)
go


CREATE NONCLUSTERED INDEX XIF1물품기초 ON mstSTK
(
	kindcd                ASC
)
go


CREATE NONCLUSTERED INDEX XIE1물품기초 ON mstSTK
(
	stknm                 ASC
)
go


CREATE NONCLUSTERED INDEX XIE2물품기초 ON mstSTK
(
	kindcd                ASC,
	stkcd                 ASC
)
go


CREATE NONCLUSTERED INDEX XIF2물품기초 ON mstSTK
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
	ADD CONSTRAINT XPK물품분류 PRIMARY KEY  CLUSTERED (kindcd ASC)
go


CREATE TABLE mstTEST
(
	testcd               varchar(7)  NOT NULL ,
	testnm               varchar(30)  NULL 
)
go


ALTER TABLE mstTEST
	ADD CONSTRAINT XPK검사항목 PRIMARY KEY  CLUSTERED (testcd ASC)
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
	ADD CONSTRAINT XPK사용자기초 PRIMARY KEY  CLUSTERED (usercd ASC)
go


CREATE NONCLUSTERED INDEX XIF1사용자기초 ON mstUSER
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
	ADD CONSTRAINT XPK장비운영내역 PRIMARY KEY  CLUSTERED (machcd ASC,operdt ASC,operseq ASC)
go


CREATE NONCLUSTERED INDEX XIF1장비운영내역 ON operL
(
	machcd                ASC
)
go


CREATE NONCLUSTERED INDEX XIF2장비운영내역 ON operL
(
	opercd                ASC
)
go


CREATE NONCLUSTERED INDEX XIF3장비운영내역 ON operL
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
	ADD CONSTRAINT XPK발주서 PRIMARY KEY  CLUSTERED (ordym ASC,ordno ASC)
go


CREATE NONCLUSTERED INDEX XIF2발주서 ON ordH
(
	usercd                ASC
)
go


CREATE NONCLUSTERED INDEX XIF1발주서 ON ordH
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
	ADD CONSTRAINT XPK발주내역 PRIMARY KEY  CLUSTERED (ordym ASC,ordno ASC,ordseq ASC)
go


CREATE NONCLUSTERED INDEX XIF1발주내역 ON ordL
(
	ordym                 ASC,
	ordno                 ASC
)
go


CREATE NONCLUSTERED INDEX XIF2발주내역 ON ordL
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
	ADD CONSTRAINT XPK출고서 PRIMARY KEY  CLUSTERED (outdt ASC,outfg ASC,outseq ASC)
go


CREATE NONCLUSTERED INDEX XIF1출고서 ON outL
(
	stkcd                 ASC
)
go


CREATE NONCLUSTERED INDEX XIF2출고서 ON outL
(
	usercd                ASC
)
go


CREATE NONCLUSTERED INDEX XIE1출고서 ON outL
(
	outdt                 ASC,
	stkcd                 ASC
)
go

CREATE NONCLUSTERED INDEX XIF3출고서 ON outL
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
	ADD CONSTRAINT XPK구매요청서 PRIMARY KEY  CLUSTERED (dutycd ASC,reqdt ASC,reqseq ASC)
go


CREATE NONCLUSTERED INDEX XIF1구매요청서 ON reqL
(
	dutycd                ASC
)
go


CREATE NONCLUSTERED INDEX XIF2구매요청서 ON reqL
(
	stkcd                 ASC
)
go


CREATE NONCLUSTERED INDEX XIF3구매요청서 ON reqL
(
	ordym                 ASC,
	ordno                 ASC,
	ordseq                ASC
)
go


CREATE NONCLUSTERED INDEX XIF4구매요청서 ON reqL
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
	ADD CONSTRAINT XPK물품재고 PRIMARY KEY  CLUSTERED (stkcd ASC,rmdym ASC)
go


CREATE NONCLUSTERED INDEX XIF1물품재고 ON stkRMD
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
	ADD CONSTRAINT XPK검사별시약기초 PRIMARY KEY  CLUSTERED (testcd ASC,stkcd ASC)
go


CREATE NONCLUSTERED INDEX XIF1검사별시약기초 ON testSTK
(
	testcd                ASC
)
go


CREATE NONCLUSTERED INDEX XIF2검사별시약기초 ON testSTK
(
	stkcd                 ASC
)
go

