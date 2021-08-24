 select   p.patid
,p.patname
,p.sex
,w.deptcode
,w.acptdate
,w.orderdate
,w.slipcode
,w.ioflag
,w.spcmno
 ,p.prsnidpre||prsnidpost prsnid
 from registinfos w
 ,  resultofnum r
 ,  patmst p
 where w.patid = p.patid
 and w.spcmno = r.spcmno
 and w.ordercode = r.ordercode
  and w.acptdate between '20110608' and '20110608'
  and r.resultitemcode in ('LB2590','LC3721','XC00038','XC0021','XC0022','XC0025','XC0027','XC0028','XC0029','XC0030','XC0031','XC0032','XC0034','XC0038','XC0047','XC0048','XC0049','XC1008')
     and w.rsvacptstate < '5'
      order by w.acptdate, w.spcmno ;

select * from resultofnum
where spcmno = '1106080030';

