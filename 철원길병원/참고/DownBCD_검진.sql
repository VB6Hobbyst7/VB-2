 select res1_gum_code           
 from mresult001                
 where res1_date  = '2013-08-23' 
   and res1_gumno = '13' 
   and res1_gum_code in ('A121','H802','H805','H806','H807','H808','H809','H810','H832','H833','H834','H835','H836')
 and (res1_result is null or res1_result = '') 
