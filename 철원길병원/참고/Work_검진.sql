 select distinct                
     res1_date   AS ADT         
   , res1_gumno  AS BCD         
   , per1_name   AS PNM         
   , per1_jumin  AS JNO         
   , per1_age                   
   , per1_sex                   
   , per1_memid  AS PID         
 from mresult001 R              
    , mperson001 P              
 where res1_date  = per1_date   
   and res1_gumno = per1_gumno  
   and res1_date  between '20130823' and '20130823' 
   and res1_gum_code in ('A121','H802','H805','H806','H807','H808','H809','H810','H832','H833','H834','H835','H836')
 and (res1_result is null or res1_result = '') 
 Order by ADT, BCD 
