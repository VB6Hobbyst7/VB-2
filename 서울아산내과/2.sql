select '', a.request_date, '',a.exam_no, b.chart_no, b.person_name, '', '', b.personal_id, '0'
from meditoliss..totres a, meditoliss..total b
where a.request_date = b.request_date and a.exam_no = b.exam_no
  and a.request_date >= '20130301' and a.request_date <= '20130309' and a.result_value = ''
  and a.exam_code in ('1001','1002','1003','1004','1005','1006','1008','1009','1010','1011','1012','1013','1015','1016','1017','1018','1019','1020','1022','1024','1025','1026','1027','1031','1032','1033','1051','1052','1053','1055','L110040','L11017501','L11019501','L110235','L110325','L110545','L110905')
and b.person_name = 'Àüµ¿¼÷'
group by a.request_date, a.exam_no, b.person_name, b.personal_id,b.chart_no
order by a.exam_no


select '', a.request_date, '',a.exam_no, b.chart_no, b.person_name, '', '', b.personal_id, '2'
from meditoliss..twoexam a, meditoliss..total b, meditoliss..panjong2 c
where a.request_date = b.request_date and a.exam_no = b.exam_no
  and a.request_date = c.request_date and a.exam_no = c.exam_no
  and a.request_date >= '20010201' and a.request_date <= '20130309'
  and a.exam_code in ('1001','1002','1003','1004','1005','1006','1008','1009','1010','1011','1012','1013','1015','1016','1017','1018','1019','1020','1022','1024','1025','1026','1027','1031','1032','1033','1051','1052','1053','1055','L110040','L11017501','L11019501','L110235','L110325','L110545','L110905')
  and b.person_name = 'Àüµ¿¼÷'
group by a.request_date, a.exam_no, b.person_name, b.personal_id,b.chart_no
order by a.exam_no


select * from meditoliss..total
 where person_name = 'Àüµ¿¼÷'