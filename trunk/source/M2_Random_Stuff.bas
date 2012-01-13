Attribute VB_Name = "M2_Random_Stuff"
Sub VariableInitialConditions()
Attribute VariableInitialConditions.VB_ProcData.VB_Invoke_Func = " \n14"
   Dim random_factor As Double
            
     random_factor = Exp(Zvector(iz) * InitialCV - 0.5 * InitialCV ^ 2)
         
     For Area = 1 To Nareas
        
        Bvulnerable(StYear, Area) = 0
        Btotal(StYear, Area) = 0
        
        For age = Stage + 1 To AgePlus
        
            N(StYear, Area, age) = N(StYear, Area, age) * random_factor
            Btotal(StYear, Area) = Btotal(StYear, Area) + N(StYear, Area, age) * w(StYear, Area, age)
            Bvulnerable(StYear, Area) = Bvulnerable(StYear, Area) + N(StYear, Area, age) * w(StYear, Area, age) * FracSel(Area, age)
          
        Next age
     Next Area
     
     Call Preliminary_Calcs.Initialize_tmp_variables
  
     If PartialSurveyFlag = True Then
      
        For Area = 1 To Nareas
           Call DoSurvey(StYear, Area)
           Atlas(Area) = SurveyBvul(1, StYear, Area)
        Next Area
      
     End If

End Sub

Sub RecruitmentDevs()
Attribute RecruitmentDevs.VB_ProcData.VB_Invoke_Func = " \n14"
   'esto es para autocorrelacionados en el tiempo pero no en el espacio
   'hay que generalizar
    Dim year As Integer
     
    For Area = 1 To Nareas
       iz = iz + 1
       Rdev(StYear, Area) = RecCV * Zvector(iz)
          
       For year = StYear + 1 To EndYear
          iz = iz + 1
          Rdev(year, Area) = (1 - RecTimeCor ^ 2) ^ 0.5 * RecCV * Zvector(iz) + RecTimeCor * Rdev(year - 1, Area)
       Next year
     
    Next Area
     
End Sub
