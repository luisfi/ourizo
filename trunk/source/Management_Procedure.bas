Attribute VB_Name = "Management_Procedure"
Dim ExpectedCatch As Double, FishedSurface As Double
Dim Nfishedareas As Integer, AtlasALL As Double, rr As Integer, i_area As Integer, i_survey
Dim Reopen As Boolean, HasReOpenConditions As Boolean, ShortenPeriod As Boolean, tempbio As Double

Sub DoSurvey(year, Area)
Attribute DoSurvey.VB_ProcData.VB_Invoke_Func = " \n14"

'Van al input

For i_survey = 1 To Nsurveys
    
 '   Select Case SurveyUnit
    
  '  Case 1 ' Numbers
    
   '  MsgBox ("Need to provide code in Management Procedure Module")
    ' End
    'Case 2 ' Biomass
     '   Select Case SurveyVariable
        
      '  Case 1 ' Bvul
          
                ' Para las geoducks tambien un error de misespeficicacion por asumir que las areas no survey estan igual que cuando surveyed (solo _
                una fraccion es surveyed each year con el supuesto que las otras areas presurveyed no cambian entre surveys excepto por pesca)
                iz = iz + 1
                SurveyBvul(i_survey, year, Area) = Bvulnerable(year, Area) * Exp(Zvector(iz) * SurveyCV - 0.5 * SurveyCV ^ 2)
                ' Aqui aun no esta calculada Bmature(year, Area), así que la comentamos:
                ' SurveyMat(i_survey, year, Area) = Bmature(year, Area) * Exp(Zvector(iz) * SurveyCV - 0.5 * SurveyCV ^ 2)
                SurveyNtot(i_survey, year, Area) = 0
                For age = 1 To Nages
                    SurveyNtot(i_survey, year, Area) = SurveyNtot(i_survey, year, Area) + n(year, Area, age)
                Next age
                SurveyNtot(i_survey, year, Area) = SurveyNtot(i_survey, year, Area) * Exp(Zvector(iz) * SurveyCV - 0.5 * SurveyCV ^ 2)
                SurveyAll(year) = SurveyAll(year) + SurveyBvul(i_survey, year, Area)
          
            
            If pLopt Then
            
            Dim temppL() As Double
            ReDim temppL(1 To Nilens)
           
            
                For i = 1 To Nilens
                    temppL(i) = pL(year, Area, i)
                Next i
            
                tempSurveypL = rmultinom(Int(SurveyNtot(i_survey, year, Area)), l, temppL)
                For i = 1 To Nilens
                    SurveypL(i_survey, year, Area, i) = tempSurveypL(i)
                Next i
           
            
            End If
        
       ' Case 2 ' Bmat
        ' MsgBox ("Need to provide code in Management Procedure Module")
         'End
                
          '  For Area = 1 To Nareas
                ' Para las geoducks tambien un error de misespeficicacion por asumir que las areas no survey estan igual que cuando surveyed (solo _
                una fraccion es surveyed each year con el supuesto que las otras areas presurveyed no cambian entre surveys excepto por pesca)
           '     iz = iz + 1
            '    Survey(i_survey, year, Area) = Bmature(year, Area) * Exp(Zvector(iz) * SurveyCV - 0.5 * SurveyCV ^ 2)
            '    SurveyAll(year) = SurveyAll(year) + Survey(i_survey, year, , Area)
            'Next Area
        
        'Case 3 ' Btot
         '   For Area = 1 To Nareas
                ' Para las geoducks tambien un error de misespeficicacion por asumir que las areas no survey estan igual que cuando surveyed (solo _
                una fraccion es surveyed each year con el supuesto que las otras areas presurveyed no cambian entre surveys excepto por pesca)
          '      iz = iz + 1
           '     Survey(i_survey, year, Area) = Btotal(year, Area) * Exp(Zvector(iz) * SurveyCV - 0.5 * SurveyCV ^ 2)
            '    SurveyAll(year) = SurveyAll(year) + Survey(i_survey, year, , Area)
            'Next Area
        'Case 4 ' User specified selectivity
         '   For Area = 1 To Nareas
                ' Para las geoducks tambien un error de misespeficicacion por asumir que las areas no survey estan igual que cuando surveyed (solo _
                una fraccion es surveyed each year con el supuesto que las otras areas presurveyed no cambian entre surveys excepto por pesca)
          '      iz = iz + 1
           '     Survey(i_survey, year, , Area) = Btotal(year, Area) * Exp(Zvector(iz) * SurveyCV - 0.5 * SurveyCV ^ 2)
           '     SurveyAll(year) = SurveyAll(year) + Survey(i_survey, year, , Area)
           ' Next Area
       ' End Select
   ' End Select
Next i_survey

End Sub
Sub CheckOpenConditions(year, Area)

               Reopen = True ' Set to True by default (as to not affect And Statement)
               HasReOpenConditions = False 'Set to False by Default
               ShortenPeriod = True 'Set to True by default (as to not affect And Statement)
               
               If ReOpenConditionFlag Then 'If there are ReOpenConditions to be evaluated
                 
                'Preharvest vulnerable biomass
                 If RCPreharvestBiomass_Fraction(Area) <> 0 Then 'If There are Preharvest Fractions set (different from 0)
                    HasReOpenConditions = True
                       ReOpenConditionValues(1, Area) = SurveyBvul(1, year, Area) / RCPreharvestBiomass_Fraction(Area) * VB0(Area)
                       Reopen = Reopen And (ReOpenConditionValues(1, Area) >= 1)
                       If AdaptativeRotationFlag Then
                        ShortenPeriod = ShortenPeriod And (ReOpenConditionValues(1, Area) >= 1 + 1 * RCPreharvestBiomass_Tolerance(Area))
                       End If
                 Else 'If Preharvest Fractions set to zero (not set)
                    ReOpenConditionValues(1, Area) = 0
                 End If
                'Minimum density threshold of individuals above the legal size
                If RCMinimumDensity(Area) <> 0 Then 'If There are Minimum Density thresholds set (different from 0)
                    HasReOpenConditions = True
                       j = l(Nilens)
                       Dim temppL As Double
                       temppL = 0
                       Do While j >= Lfull(Area)
                         temppL = SurveypL(1, year, Area, j) + temppL
                       j = j - Linc
                       Loop
                       ReOpenConditionValues(2, Area) = temppL * SurveyNtot(1, year, Area) / Surface(Area) / RCMinimumDensity(Area)
                       Reopen = Reopen And (ReOpenConditionValues(2, Area) >= 1)
                       If AdaptativeRotationFlag Then
                        ShortenPeriod = ShortenPeriod And (ReOpenConditionValues(2, Area) >= 1 + 1 * RCMinimumDensity_Tolerance(Area))
                       End If
                 Else 'If Preharvest Fractions set to zero (not set)
                    ReOpenConditionValues(2, Area) = 0
                 End If
                '%Mature Biomass
                If RCMatures_Fraction(Area) <> 0 Then 'If There are Fraction of Mature Biomass thresholds set (different from 0)
                    HasReOpenConditions = True
                       ReOpenConditionValues(3, Area) = SurveyMat(1, year, Area) / SurveyBvul(1, year, Area) / RCMatures_Fraction(Area)
                       Reopen = Reopen And (ReOpenConditionValues(3, Area) >= 1)
                       If AdaptativeRotationFlag Then
                        ShortenPeriod = ShortenPeriod And (ReOpenConditionValues(3, Area) >= 1 + 1 * RCMatures_Tolerance(Area))
                       End If
                 Else 'If Preharvest Fractions set to zero (not set)
                    ReOpenConditionValues(3, Area) = 0
                 End If
                
                '%Biomass greater than XSize
                If RCGreaterSize_Fraction(Area) <> 0 Then 'If There are Fraction of Biomass greater than X in size thresholds set (different from 0)
                    HasReOpenConditions = True
                    tempbio = 0
                    For ilen = Int((RCGreaterSize_Size(Area) - L1) / Linc + 1) To Nilens
                    tempbio = SurveyNtot(1, year, Area) * SurveypL(1, year, Area, ilen) + tempbio 'Number of individuals of size greater than X
                    Next ilen
                       ReOpenConditionValues(4, Area) = tempbio / SurveyNtot(1, year, Area) / RCGreaterSize_Fraction(Area)
                       Reopen = Reopen And (ReOpenConditionValues(4, Area) >= 1)
                       If AdaptativeRotationFlag Then
                        ShortenPeriod = ShortenPeriod And (ReOpenConditionValues(4, Area) >= 1 + 1 * RCGreaterSize_Tolerance(Area))
                       End If
                 Else 'If Preharvest Fractions set to zero (not set)
                    ReOpenConditionValues(4, Area) = 0
                 End If
                 
               'End Reopening Conditions
               End If
               
End Sub

Sub Strategies(year)
Attribute Strategies.VB_ProcData.VB_Invoke_Func = " \n14"
        
    Select Case RunFlags.Hstrategy
    
    
    
    Case 1     'ROTATION
    'Todo lo ANUAL que cierra y abre areas va a aca. y pasa a fishing un tmp que dice que areas van a ser pescadas.
    
       'Debug.Print Candidate_areas(1, 1)
   
     
        For Area = 1 To Nareas
            ClosedArea(year, Area) = True
        Next Area
        
      Select Case RunFlags.RotationType
            
      Case 1 'Rotation by Global TAC: Geoduck case study
      
        If Feedback = True Then
            If PartialSurveyFlag = True Then
                AtlasALL = 0
                For Area = 1 To Nareas
                    AtlasALL = AtlasALL + Atlas(Area)
                Next Area
                TAC(year) = AtlasALL * TargetHR
            Else
                For Area = 1 To Nareas
                    Call DoSurvey(year, Area)
                Next Area
                TAC(year) = SurveyAll(year) * TargetHR
            End If
        End If
                         
          ExpectedCatch = 0
          Nfishedareas = 0
          Do While (Nfishedareas <= Nareas) 'NB! it can enter these loops more than once until condition is satisfied
             For rr = 1 To Nregions
                For i = 1 To Nareas_region(rr)
                   Area = Candidate_areas(rr, i)
                        
                   If PartialSurveyFlag = True Then
                      iz = iz + 1
                      SurveyBvul(1, year, Area) = Bvulnerable(year, Area) * Exp(Zvector(iz) * SurveyCV - 0.5 * SurveyCV ^ 2)
                      Atlas(Area) = SurveyBvul(1, year, Area)
                   End If
                        
                   If (ReOpenConditionFlag = False) Or (SurveyBvul(1, year, Area) >= ReOpenCondition(1) * VB0(Area)) Then
                                
                      ClosedArea(year, Area) = False
   
                      Nfishedareas = Nfishedareas + 1
                      ExpectedCatch = ExpectedCatch + SurveyBvul(1, year, Area) * PulseHR 'podriamos querer reemplazar PulseHR como una densidad umbral como fraccion de capacidad de Carga
                                                       
                      For j = i To Nareas_region(rr) - 1
                          Candidate_areas(rr, j) = Candidate_areas(rr, j + 1)
                      Next j
                      
                      Candidate_areas(rr, Nareas_region(rr)) = Area
                                                        
                      Exit For 'go to next region
                                               
                   End If
                        
                Next i
                
                If (ExpectedCatch >= 0.95 * TAC(year)) Then Exit Do
             Next rr
          Loop
            
          PulseHRadjust = 1
          If (ExpectedCatch >= 1.05 * TAC(year)) Then      'need to reduce PulseHR so that TAC is not exceeded when chosen areas are opened
                    PulseHRadjust = (1.05 * TAC(year)) / ExpectedCatch
          End If
      Case 2 'Rotation by TAE
         
         MsgBox ("Rotation scheme not implemented for total effort (input) control")
         End
      
      Case 3 ' Rotation by SurfaceArea: harvest rate implemented based on surface area
         
         FishedSurface = 0
         Nfishedareas = 0
      
         Do While (FishedSurface < TargetSurface) And (Nfishedareas <= Nareas)
            For rr = 1 To Nregions
               For i = 1 To Nareas_region(rr)
                  Area = Candidate_areas(rr, i)
                  SurveyBvul(1, year, Area) = Bvulnerable(year, Area) * Exp(normal(0, SurveyCV))
                  
                  If (ReOpenConditionFlag = False) Or (SurveyBvul(1, year, Area) >= ReOpenCondition(1) * VB0(Area)) Then
                    
                       ClosedArea(year, Area) = False
                       Nfishedareas = Nfishedareas + 1
                       FishedSurface = FishedSurface + Surface(Area)
                       For j = i To Nareas_region(rr) - 1
                             Candidate_areas(rr, j) = Candidate_areas(rr, j + 1)
                       Next j
                       Candidate_areas(rr, Nareas_region(rr)) = Area
                       Exit For 'go to next region
                  End If
               Next i
               
               If (FishedSurface >= 0.9 * TargetSurface) Then Exit Do      'this strategy makes sense only if areas are small, otherwise TargetSurface will be easily exceeded
            Next rr
         Loop
      
      Case 4 ' Rotation by Period
      
         If Feedback = True Then
            If PartialSurveyFlag = True Then
                For Area = 1 To Nareas
                   If (RestingTime(Area) >= RotationPeriod(Area)) Then
                    Call DoSurvey(year, Area)
                   End If
                Next Area
            Else ' Survey all areas and compute TACs
                For Area = 1 To Nareas
                    Call DoSurvey(year, Area)
                    TAC_area(year, Area) = SurveyBvul(1, year, Area) * TargetHR
                Next Area
                    TAC(year) = SurveyAll(year) * TargetHR
            End If
         End If
       
                   
            Nfishedareas = 0
            For Area = 1 To Nareas
            
             Call CheckOpenConditions(year, Area)
            
             'Open area if RestingTime equals or exceeds rotation period and reopen conditions are met (when needed).
             If (RestingTime(Area) >= RotationPeriod(Area)) And ((ReOpenConditionFlag = False) Or (HasReOpenConditions And Reopen)) Then
                ClosedArea(year, Area) = False
                Nfishedareas = Nfishedareas + 1
                If HasReOpenConditions And ShortenPeriod And AdaptativeRotationFlag Then 'If there are ReOpenConditions and criteria for shortening rotation period is met
                    RotationPeriod(Area) = RotationPeriod(Area) - 1 'Shorten rotation period
                End If
                RestingTime(Area) = 1 'Reset resting time
             Else
                RestingTime(Area) = RestingTime(Area) + 1
                If AdaptativeRotationFlag And (RestingTime(Area) >= RotationPeriod(Area)) Then
                   RotationPeriod(Area) = RotationPeriod(Area) + 1
                End If
             End If
            Next Area
          
          
                            
      End Select

    Case 2    'AREA BY AREA MANAGEMENT - ANUAL
        
        If Feedback = True Then
           
           For Area = 1 To Nareas
            Call DoSurvey(year, Area)
           Next Area
        
           If TAC_TAE_HR = 1 Then
                For Area = 1 To Nareas
                    TAC_area(year, Area) = SurveyBvul(1, year, Area) * TargetHR
                Next Area
           ElseIf TAC_TAE_HR = 2 Then
                MsgBox ("For chosen MP you need to implement a feedback rule for effort to calculate TAE_area(Year,area)")
                End   'or end?
           End If
        End If
  
    Case 3    'Management by region
      
        If Feedback = True Then
            For Area = 1 To Nareas
                Call DoSurvey(year, Area)
            Next Area
           
            If TAC_TAE_HR = 1 Then
             
                    For rr = 1 To Nregions
                         For i_area = 1 To Nareas_region(rr)
                             Area = Candidate_areas(rr, i_area)
                             TAC_region(rr, year) = TAC_region(rr, year) + SurveyBvul(1, year, Area) * TargetHR
                         Next i_area
                    Next rr
            
            ElseIf TAC_TAE_HR = 2 Then
                    MsgBox ("For chosen MP you need to implement a feedback rule for effort to calculate TAE_region(Year,area)")
                    End   'or end?
            End If
        End If
        
        
        If TAC_TAE_HR = 1 Then
           
           EffortPulse = MaxEffort / (Nt_Season * Npulses)

        ElseIf TAC_TAE_HR = 2 Then
        ReDim EffortPulseRegion(Nregions)
           
           For rr = 1 To Nregions
              EffortPulseRegion(rr) = TAE_region(rr, year) / (Nt_Season * Npulses)
           Next rr
        
        End If
      
      
        'CR_all = 0
        'If EffortDistributionFlag = 3 Then ''Assuming allocate effort once per time step (Nt, e.g. month) using simil gravitational
      
       '     For i_area = 1 To Nopenareas
       '         area = IDopenarea(i_area)
       '         CR(area) = BvulTmp(area) * Q(area)
       '         CR_all = CR_all + CR(area)
       '
                
       '     Next i_area
       '  End If
   
   
   
    End Select

'THIS NEEDS TO BE SET INDEPENDENT OF HARVESTING STRATEGY

   For Area = 1 To Nareas
       ClosedAreaTmp(Area) = ClosedArea(year, Area)
   Next Area

End Sub
