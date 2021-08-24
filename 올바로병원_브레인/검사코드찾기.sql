SELECT DISTINCT        SLABW_INOUT             AS INOUT    
     , SLABWS_DATE             AS HOSPDATE 
     , CHAM_KEY                AS BARCODE  
,SLABW_TIME
, CHAM_WHANJA             AS PNAME    
     , CHAM_JUMIN1             AS JUMIN1   
     , CHAM_JUMIN2             AS JUMIN2   
     , SPECI_DATE              AS CHARTNO 
     , SPECI_SEQNO             AS PID      
     , CONCAT(RTRIM(LTRIM(C.SLABWS_MOMU)),'|',RTRIM(LTRIM(C.SLABWS_SCNT))) AS ITEM               FROM BRWONMU..WCHAM A                                                
       INNER JOIN OSLABW B     ON A.CHAM_KEY = B.SLABW_CHAM            
       INNER JOIN OSLABWS C    ON B.SLABW_DATE = C.SLABWS_DATE         
                               AND B.slabw_dept = C.slabws_dept        
                               AND B.slabw_cnt = C.slabws_cnt          
                               AND B.slabw_slab = C.slabws_slab        
       INNER JOIN OSLABS E     ON C.SLABWS_SCNT = E.SLABS_CNT          
                               AND C.slabws_slab = E.slabs_key         
                               AND E.slabs_use  = 1                    
       INNER JOIN Ospecislab F ON B.slabw_cnt = F.specis_cnt           
                               AND B.slabw_date = F.specis_date        
                               AND B.slabw_dept = F.specis_dept        
                               AND F.specis_deleted = 0                
       INNER JOIN OSPECIMEN S  ON A.cham_key = S.SPECI_CHAM            
                               AND F.specis_date = S.speci_date        
                               AND F.specis_seqno = S.speci_seqno      
                               And F.specis_date  = '20210121'  
                               And F.specis_seqno  = '69'  

