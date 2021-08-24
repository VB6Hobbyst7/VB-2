 SELECT j011m.bcno AS bcno, j010m.bcprtno AS prtno      
       , r010m.WKYMD||r010m.WKGRPCD||r010m.WKNO FLWKNO  
       , r010m.WKNO                                     
       , j011m.regno                                    
       , j010m.patnm                                    
       , j010m.age                                      
       , j010m.sex                                      
       , j011m.IOGBN                                    
       , j010m.DEPTCD                                   
       , j010m.WARDNO                                   
       , j010m.ROOMNO                                   
  FROM LJ011M j011m                                     
       INNER JOIN LJ010M j010m                          
               ON j011m.bcno  = j010m.bcno              
              AND j011m.regno = j010m.regno             
       INNER JOIN LR010M r010m                          
               ON j011m.bcno   = r010m.bcno             
              AND j011m.regno  = r010m.regno            
              AND NVL(r010m.rstflg,'0') = '0'       
       INNER JOIN LF072M f72m                           
               ON f72m.eqcd    = 'G0006' 
              AND f72m.testcd  = '81210'   
              AND r010m.testcd = f72m.testcd            
 WHERE j011m.colldt BETWEEN '20160310000000' and '20160310235959'  
   and r010m.wkno between '000001' and '009999' 
   AND j011m.spcflg  = '4'                        
   AND NVL(j011m.rstflg, '0')  = '0'            
 UNION                                              
 SELECT j011m.bcno AS bcno, j010m.bcprtno AS prtno  
        , r010m.FLWKNO                              
        , r010m.WKNO                                
        , j011m.regno                               
        , j010m.patnm                               
        , j010m.age                                 
        , j010m.sex                                 
        , j011m.IOGBN                               
        , j010m.DEPTCD                              
        , j010m.WARDNO                              
        , j010m.ROOMNO                              
   FROM LJ011M j011m                                
        INNER JOIN LJ010M j010m                     
                ON j011m.bcno  = j010m.bcno         
               AND j011m.regno = j010m.regno        
        INNER JOIN LM010M r010m                     
                ON j011m.bcno   = r010m.bcno        
               AND j011m.regno  = r010m.regno       
               AND NVL(r010m.rstflg,'0') = '0'  
        INNER JOIN LF072M f72m                      
                ON f72m.eqcd    = 'G0006' 
                AND f72m.testcd  = '81210'  
               AND r010m.testcd = f72m.testcd       
  WHERE j011m.colldt BETWEEN '20160310000000' and '20160310235959'  
   and r010m.wkno between'000001' and '009999' 
    AND j011m.spcflg  = '4'               
    AND NVL(j011m.rstflg, '0')  = '0'     
    ORDER BY FLWKNO  
