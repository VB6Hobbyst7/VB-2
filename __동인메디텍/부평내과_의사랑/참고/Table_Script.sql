DB Name : DDIF

CREATE TABLE dbo.calculation (
    IFCC1    VARCHAR(5), 
    IFCC2    VARCHAR(5), 
    EAG1     VARCHAR(5), 
    EAG2     VARCHAR(5), 
    ADDIFCC  VARCHAR(1), 
    ADDEAG   VARCHAR(1)
) ON [PRIMARY]


CREATE TABLE dbo.equipexam (
    equipno     VARCHAR(20), 
    equipcode   VARCHAR(15), 
    examcode    VARCHAR(15), 
    examname    VARCHAR(20), 
    resprec     SMALLINT DEFAULT (1), 
    reflow      VARCHAR(20), 
    refhigh     VARCHAR(20), 
    paniclow    VARCHAR(20), 
    panichigh   VARCHAR(20), 
    deltavalue  VARCHAR(50), 
    seqno       SMALLINT DEFAULT (0), 
    examflag    SMALLINT DEFAULT (1), 
    examtype    VARCHAR(50)
) ON [PRIMARY]


CREATE TABLE dbo.pat_res (
    Company    VARCHAR(8), 
    HospCode   VARCHAR(20), 
    ChartNo    VARCHAR(20), 
    PatName    VARCHAR(20), 
    PatSex     VARCHAR(1), 
    PatAge     VARCHAR(5), 
    PatJumin   VARCHAR(20), 
    PatNo      VARCHAR(50), 
    CommDate   VARCHAR(20), 
    ExamNo     VARCHAR(20), 
    ExamID     VARCHAR(20), 
    ComExamID  VARCHAR(20), 
    Specimen   VARCHAR(20), 
    Result     VARCHAR(20), 
    Reference  VARCHAR(30), 
    Remark     VARCHAR(20), 
    RsltDate   VARCHAR(8), 
    IOFlag     VARCHAR(1), 
    TransYN    VARCHAR(1), 
    TransDT    VARCHAR(8), 
    Barcode    VARCHAR(12), 
    examtype   VARCHAR(50)
) ON [PRIMARY]



CREATE TABLE dbo.qc_res (
    equipno    VARCHAR(20), 
    examdate   VARCHAR(8), 
    examtime   VARCHAR(10), 
    levelname  VARCHAR(20), 
    equipcode  VARCHAR(15), 
    result     VARCHAR(20), 
    resflag    VARCHAR(10), 
    remark     VARCHAR(50), 
    examuid    VARCHAR(10), 
    sresult    VARCHAR(10), 
    lotno      VARCHAR(20)
) ON [PRIMARY]


CREATE TABLE dbo.qcexam (
    equipno     VARCHAR(20), 
    lotno       VARCHAR(20), 
    levelno     SMALLINT, 
    levelname   VARCHAR(10), 
    appdate     VARCHAR(8), 
    validstart  VARCHAR(8), 
    validend    VARCHAR(8), 
    equipcode   VARCHAR(15), 
    examname    VARCHAR(20), 
    t_mean      VARCHAR(10), 
    t_sd        VARCHAR(10), 
    remark      VARCHAR(50)
) ON [PRIMARY]
