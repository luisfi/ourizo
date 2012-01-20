Attribute VB_Name = "Management_Procedure"
Dim ExpectedCatch As Double, FishedSurface As Double
Dim Nfishedareas As Integer, AtlasALL As Double, rr As Integer, i_area As Integer, i_survey
Dim reopen As Boolean, HasReOpenConditions As Boolean, ShortenPeriod As Boolean, tempbio As Double

Sub DoSurvey(year, area)
Attribute DoSurvey.VB_ProcData.VB_Invoke_Func = " \n14"

'Van al input

  For i_survey = 1 To Nsurveys
    ' Para las geoducks tambien un error de misespeficicacion por asumir que las areas no survey estan igual que cuando surveyed (solo _
    una fraccion es surveyed each year con el supuesto que las otras areas presurveyed no cambian entre surveys excepto por pesca)
    iz = iz + 1
    'SurveyBvul(i_survey, year, Area) = Bvulnerable(year, Area) * Exp(Zvector(iz) * SurveyCV - 0.5 * SurveyCV ^ 2)
    'SurveyBtot(i_survey, year, Area) = Btotal(year, Area) * Exp(Zvector(iz) * SurveyCV - 0.5 * SurveyCV ^ 2)
    
    SurveyNtot(i_survey, year, area) = 0
    SurveyBvul(i_survey, year, area) = 0
    SurveyBtot(i_survey, year, area) = 0
    
    For age = 1 To Nages
        SurveyNtot(i_survey, year, area) = SurveyNtot(i_survey, year, area) + N(year, area, age)
    Next age
    SurveyNtot(i_survey, year, area) = SurveyNtot(i_survey, year, area) * Exp(Zvector(iz) * SurveyCV - 0.5 * SurveyCV ^ 2)
            
     If sample_size_pL > 0 Then
     
     Dim temppL() As Double
     ReDim temppL(1 To Nilens)
   
         For i = 1 To Nilens
             temppL(i) = pL(year, area, i)
         Next i
         
         tempSurveypL = rmultinom(sample_size_pL, l, temppL)
         
         For i = 1 To Nilens
             SurveypL(i_survey, year, area, i) = tempSurveypL(i)
             SurveyBtot(i_survey, year, area) = SurveyBtot(i_survey, year, area) + SurveyNtot(i_survey, year, area) * SurveypL(i_survey, year, area, i) * W_L(area, i)
             If l(i) >= Lfull(area) Then SurveyBvul(i_survey, year, area) = SurveyBvul(i_survey, year, area) + SurveyNtot(i_survey, year, area) * SurveypL(i_survey, year, area, i) * W_L(area, i)
         Next i
     
     End If

  Next i_survey

End Sub
Sub CheckOpenConditions(year, area)

               'Reopen = True ' Set to True by default (as to not affect And Statement)
               HasReOpenConditions = False 'Set to False by Default
               ShortenPeriod = True 'Set to True by default (as to not affect And Statement)
               
               If ReOpenConditionFlag Then 'If there are ReOpenConditions to be evaluated
                 
                'Virgin vulnerable biomass fraction
                 If RCVirginBiomass_Fraction(area) <> 0 Then 'If There are Virgin Fractions set (different from 0)
                    HasReOpenConditions = True
                       ReOpenConditionValues(1, area) = SurveyBvul(1, year, area) / (RCVirginBiomass_Fraction(area) * VBvirgin(area))
                       reopen = reopen And (ReOpenConditionValues(1, area) >= 1)
                       If AdaptativeRotationFlag Then
                        ShortenPeriod = ShortenPeriod And (ReOpenConditionValues(1, area) >= 1 + RCVirginBiomass_Tolerance(area))
                       End If
                 Else 'If Preharvest Fractions set to zero (not set)
                    ReOpenConditionValues(1, area) = 0
                 End If
                'Preharvest vulnerable biomass fraction
                 If RCPreharvestBiomass_Fraction(area) <> 0 Then
                   HasReOpenConditions = True
                   If (year - RestingTime(area) >= StYear) Then 'If it has been harvested before in the simulation: check reopening condition.
                     ReOpenConditionValues(2, area) = SurveyBvul(1, year, area) / (RCPreharvestBiomass_Fraction(area) * SurveyBvul(1, year - RestingTime(area), area))
                     reopen = reopen And (ReOpenConditionValues(2, area) >= 1) 'Close area if reopening condition is not met.
                     If AdaptativeRotationFlag Then
                       ShortenPeriod = ShortenPeriod And (ReOpenConditionValues(2, area) >= 1 + RCPreharvestBiomass_Tolerance(area))
                     End If
                   End If
                 Else 'If Preharvest Fractions set to zero (not set)
                    ReOpenConditionValues(2, area) = 0
                 End If
                'Minimum density threshold of individuals above the legal size
                 If RCMinimumDensity(area) <> 0 Then 'If There are Minimum Density thresholds set (different from 0)
                    HasReOpenConditions = True
                       j = Nilens
                       Dim temppL As Double
                       temppL = 0
                       Do While l(j) >= Lfull(area)
                         temppL = SurveypL(1, year, area, j) + temppL
                       j = j - 1
                       Loop
                       ReOpenConditionValues(3, area) = (temppL * SurveyNtot(1, year, area) / Surface(area)) / RCMinimumDensity(area)
                       reopen = reopen And (ReOpenConditionValues(3, area) >= 1)
                       If AdaptativeRotationFlag Then
                         ShortenPeriod = ShortenPeriod And (ReOpenConditionValues(3, area) >= 1 + RCMinimumDensity_Tolerance(area))
                       End If
                 Else 'If Preharvest Fractions set to zero (not set)
                    ReOpenConditionValues(3, area) = 0
                 End If
                
                '%Biomass greater than XSize
                If RCGreaterSize_Fraction(area) <> 0 Then 'If There are Fraction of Biomass greater than X in size thresholds set (different from 0)
                    HasReOpenConditions = True
                    tempbio = 0
                    For ilen = Int((RCGreaterSize_Size(area) - L1) / Linc + 1) To Nilens
                    tempbio = SurveyNtot(1, year, area) * SurveypL(1, year, area, ilen) + tempbio 'Number of individuals of size greater than X
                    Next ilen
                       ReOpenConditionValues(4, area) = tempbio / SurveyNtot(1, year, area) / RCGreaterSize_Fraction(area)
                       reopen = reopen And (ReOpenConditionValues(4, area) >= 1)
                       If AdaptativeRotationFlag Then
                        ShortenPeriod = ShortenPeriod And (ReOpenConditionValues(4, area) >= 1 + 1 * RCGreaterSize_Tolerance(area))
                       End If
                 Else 'If Preharvest Fractions set to zero (not set)
                    ReOpenConditionValues(4, area) = 0
                 End If
              
               'End Reopening Conditions
               End If
               
End Sub

Sub Strategies(year)
Attribute Strategies.VB_ProcData.VB_Invoke_Func = " \n14"
        
    Select Case RunFlags.Hstrategy
    
    
    
    Case 1     'ROTATION
    'Todo lo ANUAL que cierra y abre areas va a aca. y pasa a fishing un tmp que dice que areas van a ser pescadas.

        For area = 1 To Nareas
            ClosedArea(year, area) = True
        Next area
        
      Select Case RunFlags.RotationType
            
      Case 1 'Rotation by Global TAC: Geoduck case study
      
        If Feedback = True Then
            If PartialSurveyFlag = True Then
                AtlasALL = 0
                For area = 1 To Nareas
                    AtlasALL = AtlasALL + Atlas(area)
                Next area
                TAC(year) = AtlasALL * TargetHR
            Else
                For area = 1 To Nareas
                    Call DoSurvey(year, area)
                Next area
            End If
        End If
                         
          ExpectedCatch = 0
          Nfishedareas = 0
          Do While (Nfishedareas <= Nareas) 'NB! it can enter these loops more than once until condition is satisfied
             For rr = 1 To Nregions
                For i = 1 To Nareas_region(rr)
                   area = Candidate_areas(rr, i)
                        
                   If PartialSurveyFlag = True Then
                      iz = iz + 1
                      Call DoSurvey(year, area)
                      Atlas(area) = SurveyBvul(1, year, area)
                   End If
                   
                   reopen = True
                   If ReOpenConditionFlag = True Then Call CheckOpenConditions(year, area)
                  
                   If reopen Then
                   
                      ClosedArea(year, area) = False
   
                      Nfishedareas = Nfishedareas + 1
                      ExpectedCatch = ExpectedCatch + SurveyBvul(1, year, area) * PulseHR 'podriamos querer reemplazar PulseHR como una densidad umbral como fraccion de capacidad de Carga
                                                       
                      For j = i To Nareas_region(rr) - 1
                          Candidate_areas(rr, j) = Candidate_areas(rr, j + 1)
                      Next j
                      
                      Candidate_areas(rr, Nareas_region(rr)) = area 'Put opened area at the bottom of vector
                                                        
                      Exit For 'go to next region
                                               
                   End If
                        
                Next i 'Area within region
                
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
                  area = Candidate_areas(rr, i)
                  
                  Call DoSurvey(year, area)
                  reopen = True
                  If ReOpenConditionFlag = True Then Call CheckOpenConditions(year, area)
                  
                  If (reopen) Then
                    
                       ClosedArea(year, area) = False
                       Nfishedareas = Nfishedareas + 1
                       FishedSurface = FishedSurface + Surface(area)
                       For j = i To Nareas_region(rr) - 1
                             Candidate_areas(rr, j) = Candidate_areas(rr, j + 1)
                       Next j
                       Candidate_areas(rr, Nareas_region(rr)) = area
                       Exit For 'go to next region
                  End If
               Next i
               
               If (FishedSurface >= 0.9 * TargetSurface) Then Exit Do      'this strategy makes sense only if areas are small, otherwise TargetSurface will be easily exceeded
            Next rr
         Loop
      
      Case 4 ' Rotation by Period
      
         If Feedback = True Then
            If PartialSurveyFlag = True Then
                For area = 1 To Nareas
                   If (RestingTime(area) >= RotationPeriod(area)) Then
                    Call DoSurvey(year, area)
                   End If
                Next area
            Else ' Survey all areas and compute TACs
                For area = 1 To Nareas
                    Call DoSurvey(year, area)
                    TAC_area(year, area) = SurveyBvul(1, year, area) * PulseHR
                Next area
            End If
         End If
       
                   
            Nfishedareas = 0
            For area = 1 To Nareas
              If (RestingTime(area) < RotationPeriod(area)) Then
                 RestingTime(area) = RestingTime(area) + 1
              Else 'Open area if RestingTime equals or exceeds rotation period and reopen conditions are met.
                 reopen = True
                 If ReOpenConditionFlag = True Then Call CheckOpenConditions(year, area)
                 
                 If reopen Then
                    ClosedArea(year, area) = False
                    Nfishedareas = Nfishedareas + 1
                    If AdaptativeRotationFlag And ShortenPeriod Then  'If there are ReOpenConditions and criteria for shortening rotation period is met
                       RotationPeriod(area) = RotationPeriod(area) - 1 'Shorten rotation period
                    End If
                    RestingTime(area) = 1 'Reset resting time
                 Else ' Did not meet reopening conditions
                    RestingTime(area) = RestingTime(area) + 1
                    If AdaptativeRotationFlag Then RotationPeriod(area) = RotationPeriod(area) + 1
                 End If
              End If
            Next area
        
      End Select
       
    ' Print Rotational Output: monte no esta definido... Como se saca...
    Call Input_Output.Print_Rotational_Output(year)

    Case 2    'AREA BY AREA MANAGEMENT - ANUAL
        
        If Feedback = True Then
           
           For area = 1 To Nareas
            Call DoSurvey(year, area)
           Next area
        
           If TAC_TAE_HR = 1 Then
                For area = 1 To Nareas
                    TAC_area(year, area) = SurveyBvul(1, year, area) * TargetHR
                Next area
           ElseIf TAC_TAE_HR = 2 Then
                MsgBox ("For chosen MP you need to implement a feedback rule for effort to calculate TAE_area(Year,area)")
                End   'or end?
           End If
        End If
  
    Case 3    'Management by region
      
        If Feedback = True Then
            For area = 1 To Nareas
                Call DoSurvey(year, area)
            Next area
           
            If TAC_TAE_HR = 1 Then
             
                    For rr = 1 To Nregions
                         For i_area = 1 To Nareas_region(rr)
                             area = Candidate_areas(rr, i_area)
                             TAC_region(rr, year) = TAC_region(rr, year) + SurveyBvul(1, year, area) * TargetHR
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

   For area = 1 To Nareas
       ClosedAreaTmp(area) = ClosedArea(year, area)
   Next area

End Sub
