 SELECT A.LabLotTestInfo 
      , B.AnalyteID      
      , B.InstrumentID   
      , B.MethodID       
      , B.ReagentID      
      , B.TemperatureID  
      , B.UnitID         
 FROM LabLotTest A INNER JOIN Test B ON A.TestId= B.TestId 
 WHERE A.LabId=506927 
   AND A.LotID='78450' 
