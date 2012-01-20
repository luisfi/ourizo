Attribute VB_Name = "Read_Input"
Option Explicit

Dim row_parameters As Integer, row_area_atributes As Integer, row_Population_Dynamics As Integer, _
    row_global_parameters As Integer, row_parameters_biol_region As Integer, row_parameters_area As Integer, _
    row_initial_conditions As Integer, row_connectivity As Integer, _
    row_management_control As Integer, row_run_options As Integer, row_catch_specification As Integer, row_effort_specification As Integer, _
    row_input_conditioning As Integer, row_reopening_conditions As Integer, row_rotation_by_period As Integer

Sub Read_Input()

row_run_options = 5
row_input_conditioning = 18
row_parameters = 30

    Nareas = Worksheets("Input").Rows(row_parameters + 1).Columns(2)
    StYear = Worksheets("Input").Rows(row_parameters + 2).Columns(2)
    EndYear = Worksheets("Input").Rows(row_parameters + 3).Columns(2)
    Nt = Worksheets("Input").Rows(row_parameters + 4).Columns(2)
    t_Repr = Worksheets("Input").Rows(row_parameters + 5).Columns(2)
    FracHRPreRepr = Worksheets("Input").Rows(row_parameters + 6).Columns(2)
    Stage = Worksheets("Input").Rows(row_parameters + 7).Columns(2)
    AgePlus = Worksheets("Input").Rows(row_parameters + 8).Columns(2)
    L1 = Worksheets("Input").Rows(row_parameters + 9).Columns(2)
    Linc = Worksheets("Input").Rows(row_parameters + 10).Columns(2)
    Nilens = Worksheets("Input").Rows(row_parameters + 11).Columns(2)
    Nyears = EndYear - StYear + 1

row_area_atributes = 43
row_Population_Dynamics = 48
row_global_parameters = 56
row_parameters_biol_region = 62
row_parameters_area = 74
row_initial_conditions = 82

row_management_control = 89
row_reopening_conditions = row_management_control + 19
row_rotation_by_period = row_reopening_conditions + 6

row_connectivity = 144
row_catch_specification = row_connectivity + Nareas + 3
row_effort_specification = row_catch_specification + Nyears + 3


'GIVE DIMENSIONS TO BIOLOGICAL DYNAMIC OBJECTS

'area_atributes
ReDim Surface(Nareas), Lat(Nareas), Lon(Nareas)
'parameters_biol_region
ReDim Bregion(Nareas), Linf(Nareas), k(Nareas), t0(Nareas), CVmu(Nareas), aW(Nareas), bW(Nareas), M(Nareas)
'parameter_area
ReDim Kcarga(Nareas), Rmax(Nareas), q(Nareas), gk(Nareas), Bthreshold(Nareas)
'initial_conditions
ReDim HR_start(Nareas)

ReDim Connect(Nareas, Nareas)

ReDim CR(Nareas)
ReDim cost(Nareas)
ReDim ObsRec(StYear To EndYear, Nareas)

'HARDWIRED MAXIMUM NUMBER OF ELEMENTS IN Zvector
N_Zvector = 100000

ReDim Zvector(N_Zvector)

    For area = 1 To Nareas
        Surface(area) = Worksheets("Input").Rows(row_area_atributes + 2).Columns(1 + area)
        Lat(area) = Worksheets("Input").Rows(row_area_atributes + 3).Columns(1 + area)
        Lon(area) = Worksheets("Input").Rows(row_area_atributes + 4).Columns(1 + area)
    Next area

    RunFlags.Rec = Worksheets("Input").Rows(row_Population_Dynamics + 1).Columns(2)
    RunFlags.Growth_type = Worksheets("Input").Rows(row_Population_Dynamics + 5).Columns(2)
    RunFlags.VirginAgePlus = Worksheets("Input").Rows(row_Population_Dynamics + 6).Columns(2)
    RunFlags.ObsError_Survey = Worksheets("Input").Rows(13).Columns(2)
    RunFlags.ProcError_Rec = Worksheets("Input").Rows(15).Columns(2)
    RunFlags.ProcError_InitConditions = Worksheets("Input").Rows(16).Columns(2)
    
    AgeFullMature = Worksheets("Input").Rows(row_global_parameters + 1).Columns(2)
    Lambda_ProdXB = Worksheets("Input").Rows(row_global_parameters + 2).Columns(2)
    handling = Worksheets("Input").Rows(row_global_parameters + 3).Columns(2)
    price = Worksheets("Input").Rows(row_global_parameters + 4).Columns(2)
        
    NBregions = Worksheets("Input").Rows(row_parameters_biol_region + 1).Columns(2)
    For area = 1 To Nareas
        Bregion(area) = Worksheets("Input").Rows(row_parameters_biol_region + 3).Columns(area + 1)
      
        Linf(area) = Worksheets("Input").Rows(row_parameters_biol_region + 4).Columns(Bregion(area) + 1)
        k(area) = Worksheets("Input").Rows(row_parameters_biol_region + 5).Columns(Bregion(area) + 1)
        t0(area) = Worksheets("Input").Rows(row_parameters_biol_region + 6).Columns(Bregion(area) + 1)
        CVmu(area) = Worksheets("Input").Rows(row_parameters_biol_region + 7).Columns(Bregion(area) + 1)
        aW(area) = Worksheets("Input").Rows(row_parameters_biol_region + 8).Columns(Bregion(area) + 1)
        bW(area) = Worksheets("Input").Rows(row_parameters_biol_region + 9).Columns(Bregion(area) + 1)
        M(area) = Worksheets("Input").Rows(row_parameters_biol_region + 10).Columns(Bregion(area) + 1)
    Next area

    For area = 1 To Nareas
        Kcarga(area) = Worksheets("Input").Rows(row_parameters_area + 1).Columns(area + 1)
          Kcarga(area) = Kcarga(area) * Surface(area)
        Rmax(area) = Worksheets("Input").Rows(row_parameters_area + 2).Columns(area + 1)
          Rmax(area) = Rmax(area) * Surface(area)
        q(area) = Worksheets("Input").Rows(row_parameters_area + 3).Columns(area + 1)
        cost(area) = Worksheets("Input").Rows(row_parameters_area + 4).Columns(area + 1)
        If RunFlags.Growth_type = 1 Then
            gk(area) = 1
        Else
            gk(area) = Worksheets("Input").Rows(row_parameters_area + 5).Columns(area + 1)
            Bthreshold(area) = Worksheets("Input").Rows(row_parameters_area + 6).Columns(area + 1)
            Bthreshold(area) = Bthreshold(area) * Kcarga(area)
        End If
    
    Next area

    RunFlags.Initial_Conditions = Worksheets("Input").Rows(row_initial_conditions + 1).Columns(2)
    For area = 1 To Nareas
        HR_start(area) = Worksheets("Input").Rows(row_initial_conditions + 2).Columns(area + 1)
    Next area

    For area = 1 To Nareas
        For i = 1 To Nareas
            Connect(area, i) = Worksheets("Input").Rows(row_connectivity + area + 1).Columns(1 + i)
        Next i
    Next area
              
'''HERE STARTS MANAGEMENT INPUT


Nregions = Worksheets("Input").Rows(row_management_control + 1).Columns(2)

    ReDim TAC(StYear To EndYear)
    ReDim TAE(StYear To EndYear)
    
    ReDim TAE_area(StYear To EndYear, Nareas)
    ReDim TAC_area(StYear To EndYear, Nareas)
    ReDim ClosedArea(StYear To EndYear, Nareas)
    ReDim ClosedAreaTmp(Nareas)
    ReDim TAC_region(Nregions, StYear To EndYear)
    ReDim TAE_region(Nregions, StYear To EndYear)
    ReDim Region(Nareas)
    ReDim Lfull(Nareas)
    ReDim iLfull(Nareas)

For area = 1 To Nareas
    Region(area) = Worksheets("Input").Rows(row_management_control + 3).Columns(area + 1)
    ClosedArea(StYear, area) = Worksheets("Input").Rows(row_management_control + 4).Columns(1 + area)
    Lfull(area) = Worksheets("Input").Rows(row_management_control + 5).Columns(1 + area)
    iLfull(area) = ((Lfull(area) - L1) / Linc) + 1
Next area


RunFlags.Hstrategy = Worksheets("Input").Rows(row_management_control + 6).Columns(2)
TAC_TAE_HR = Worksheets("Input").Rows(row_management_control + 7).Columns(2)
MaxEffort = Worksheets("Input").Rows(row_management_control + 8).Columns(2)
Feedback = Worksheets("Input").Rows(row_management_control + 9).Columns(2)
TargetHR = Worksheets("Input").Rows(row_management_control + 10).Columns(2)


Nt_Season = Worksheets("Input").Rows(row_management_control + 11).Columns(2)
t_StSeason = Worksheets("Input").Rows(row_management_control + 12).Columns(2)
Nsurveys = Worksheets("Input").Rows(21 + row_reopening_conditions).Columns(2)


If (Nt_Season > Nt) Or (t_StSeason > Nt) Then
    MsgBox ("Season length (Nt_Season > Nt) or (t_StSeason > Nt)")
    End
End If

Select Case RunFlags.Hstrategy

Case 1  'Rotation
    
    RunFlags.RotationType = Worksheets("Input").Rows(row_management_control + 14).Columns(2)
    
    If (Nt_Season > 1) Then
        MsgBox ("Rotation is not implemented for Nt_Season> 1 ")
        End
    End If
    
    For i = StYear To EndYear
        TAC(i) = Worksheets("Input").Rows(row_catch_specification + 2 + (i - StYear)).Columns(2)
        For area = 1 To Nareas
            TAC_area(i, area) = Worksheets("Input").Rows(row_catch_specification + 2 + (i - StYear)).Columns(2 + Nregions + area)
        Next area
    Next i
        
    PartialSurveyFlag = Worksheets("Input").Rows(row_management_control + 15).Columns(2)
    
    'if there is no feedback you can't have a partial survey
    If (Feedback = False And PartialSurveyFlag = True) Then
      MsgBox ("Inconsistent flags: setting PartialSurveyFlag to False")
      PartialSurveyFlag = False
      Worksheets("Input").Rows(row_management_control + 15).Columns(2) = False
    End If
    
    PulseHR = Worksheets("Input").Rows(row_management_control + 16).Columns(2)
    
    ReOpenConditionFlag = Worksheets("Input").Rows(row_management_control + 18).Columns(2)
    'if there is no feedback you can't check ReOpenConditions
    If (Feedback = False And ReOpenConditionFlag = True) Then
      MsgBox ("Inconsistent flags: setting ReOpenConditionFlag to False")
      ReOpenConditionFlag = False
      Worksheets("Input").Rows(row_management_control + 18).Columns(2) = False
    End If
    
    'Reopen Conditions
    NOpenConditions = 4
    
    ReDim ReOpenCondition(Nareas)
    ReDim ShortenTolerance(Nareas)
    ReDim ReOpenConditionValues(Nareas, NOpenConditions)
    
    ReDim RCVirginBiomass_Fraction(Nareas)
    ReDim RCVirginBiomass_Tolerance(Nareas)
    ReDim RCPreharvestBiomass_Fraction(Nareas)
    ReDim RCPreharvestBiomass_Tolerance(Nareas)
    ReDim RCMinimumDensity(Nareas)
    ReDim RCMinimumDensity_Tolerance(Nareas)
    ReDim RCGreaterSize_Fraction(Nareas)
    ReDim RCGreaterSize_Size(Nareas)
    ReDim RCGreaterSize_Tolerance(Nareas)
    
    For i = 1 To Nareas
        RCVirginBiomass_Fraction(i) = Worksheets("Input").Rows(row_reopening_conditions + 1).Columns(i + 1)
        RCPreharvestBiomass_Fraction(i) = Worksheets("Input").Rows(row_reopening_conditions + 2).Columns(i + 1)
        RCMinimumDensity(i) = Worksheets("Input").Rows(row_reopening_conditions + 3).Columns(i + 1)
        RCGreaterSize_Fraction(i) = Worksheets("Input").Rows(row_reopening_conditions + 4).Columns(i + 1)
        RCGreaterSize_Size(i) = Worksheets("Input").Rows(row_reopening_conditions + 5).Columns(i + 1)
    Next i
    
    
    RestingTimeFlag = Worksheets("Input").Rows(row_rotation_by_period + 1).Columns(2)
    
    
    If RunFlags.Hstrategy = 1 Then 'If Rotational Strategy.
        ReDim RestingTime(Nareas)
        ReDim RotationPeriod(Nareas)
        For area = 1 To Nareas
            RestingTime(area) = Worksheets("Input").Rows(row_management_control + 17).Columns(1 + area)
            RotationPeriod(area) = Worksheets("Input").Rows(row_rotation_by_period + 2).Columns(1 + area)
        Next area
    End If
    
    AdaptativeRotationFlag = Worksheets("Input").Rows(row_rotation_by_period + 3).Columns(2)
    'if there is no ReOpenConditions you can't have and AdaptativeRotationFlag
    If (ReOpenConditionFlag = False And AdaptativeRotationFlag = True) Then
      MsgBox ("Inconsistent flags: setting AdaptativeRotationFlag to False")
      AdaptativeRotationFlag = False
      Worksheets("Input").Rows(row_rotation_by_period + 3).Columns(2) = False
    End If
    
    
    For i = 1 To Nareas
        RCVirginBiomass_Tolerance(i) = Worksheets("Input").Rows(row_rotation_by_period + 4).Columns(i + 1)
        RCPreharvestBiomass_Tolerance(i) = Worksheets("Input").Rows(row_rotation_by_period + 5).Columns(i + 1)
        RCMinimumDensity_Tolerance(i) = Worksheets("Input").Rows(row_rotation_by_period + 6).Columns(i + 1)
        RCGreaterSize_Tolerance(i) = Worksheets("Input").Rows(row_rotation_by_period + 7).Columns(i + 1)
    Next i
    
Case 2  'By Area
    If (Nt_Season > 1) Then
        MsgBox ("Area_by_area strategy not implemented for Nt_Season> 1 ")
        End
    End If

    If Feedback = False Then
            'Public TAC_area y Effort
            If TAC_TAE_HR = 1 Then
                For i = StYear To EndYear
                    For area = 1 To Nareas
                         TAC_area(i, area) = Worksheets("Input").Rows(row_catch_specification + 2 + (i - StYear)).Columns(2 + Nregions + area)
                    Next area
                Next i
            
            ElseIf TAC_TAE_HR = 2 Then
                
                For i = StYear To EndYear
                    For area = 1 To Nareas
                        TAE_area(i, area) = Worksheets("Input").Rows(row_effort_specification + 2 + (i - StYear)).Columns(2 + Nregions + area)
                    Next area
                Next i
            End If
    'Else
       '     TargetHR = Worksheets("Input").Rows(row_management_control + 10).Columns(2)
    End If


Case 3  'By region
    
    If TAC_TAE_HR = 3 Then
     
        MsgBox ("Set RunFlags.Hstrategy to 2 (management by area) if you want to simulate with known HR")
        Worksheets("Input").Activate
        End
                
    End If
    
    'Here subcase IFD or GRAVITATIONAL
    
    EffortDistributionFlag = Worksheets("Input").Rows(row_reopening_conditions + 15).Columns(2)
        
    Select Case EffortDistributionFlag
    
    Case 1  'obsolete IFD
        Sens = Worksheets("Input").Rows(row_reopening_conditions + 16).Columns(2)
        Ndias_beforeswitch = Worksheets("Input").Rows(row_reopening_conditions + 17).Columns(2)
        Npulses = 365 \ (Nt * Ndias_beforeswitch) 'integer division
    
    Case 2
        
        Npulses = 1 'Assuming allocate effort once per time step (e.g. month) using algorithm modified from Walters & MArtel
    
    Case 3
    
        '###########################################################
        'THIS OPTION NOT IMPLEMENTED YET
        MsgBox ("Option not implemented yet, change fishing effort distribution to 2(IFD)")
        Worksheets("Input").Activate
        
        End
        
    End Select
    
    If Feedback = False Then
      If TAC_TAE_HR = 1 Then
        For i = StYear To EndYear
           For j = 1 To Nregions
               TAC_region(j, i) = Worksheets("Input").Rows(row_catch_specification + 2 + (i - StYear)).Columns(2 + j)
           Next j
        Next i
      ElseIf TAC_TAE_HR = 2 Then
        For i = StYear To EndYear
           For j = 1 To Nregions
               TAE_region(j, i) = Worksheets("Input").Rows(row_effort_specification + 2 + (i - StYear)).Columns(2 + j)
           Next j
        Next i
     
      End If
    End If
End Select

RunFlags.Output_NAge_NSize = Worksheets("Input").Rows(10).Columns(2)
RunFlags.Output_Size_W = Worksheets("Input").Rows(9).Columns(2)

RunFlags.Run_type = Worksheets("Input").Rows(row_run_options + 1).Columns(2)

RunFlags.InputRec = Worksheets("Input").Rows(row_input_conditioning + 1).Columns(2)
q_Rec = Worksheets("Input").Rows(row_input_conditioning + 1).Columns(3)
RunFlags.InputBvul = Worksheets("Input").Rows(row_input_conditioning + 2).Columns(2)
RunFlags.BvulType = Worksheets("Input").Rows(row_input_conditioning + 2).Columns(3)
RunFlags.InputAbundance = Worksheets("Input").Rows(row_input_conditioning + 3).Columns(2)
RunFlags.AbundanceType = Worksheets("Input").Rows(row_input_conditioning + 3).Columns(3)
RunFlags.InputCatch = Worksheets("Input").Rows(row_input_conditioning + 4).Columns(2)

sample_size_pL = Worksheets("Input").Rows(row_rotation_by_period + 16 + 1).Columns(2)
SurveyCV = Worksheets("Input").Rows(row_rotation_by_period + 16 + 3).Columns(2)
InitialCV = Worksheets("Input").Rows(row_initial_conditions + 3).Columns(2)
RecCV = Worksheets("Input").Rows(row_Population_Dynamics + 2).Columns(2)

'Read Input Data for conditioning
If RunFlags.Run_type = 1 Then
    Nreplicates = 1
    
    If RunFlags.InputRec = True Then
        'read Observed Recruitment
        'q_rec = READ FROM INPUT
        
        For i = StYear To EndYear
            For area = 1 To Nareas
                ObsRec(i, area) = Worksheets("ObsRec").Rows((i - StYear + 1 + Nyears * (area - 1))).Columns(1)
            Next area
        Next i
    End If
    
    If RunFlags.InputBvul = True Then
        'read Observed Bvulnerable
        i = 0
        Do
            i = i + 1
            j = Worksheets("ObsBvul").Rows(i + 1).Columns(1)
        Loop While (j = Empty) = False
        If i = 1 Then MsgBox ("Include Bvul Data or change Input Bvul Flag")
        NObsBvul = i - 1
        
        ReDim ObsBvul(NObsBvul, 4)
        
        For i = 1 To NObsBvul
            For j = 1 To 4
                ObsBvul(i, j) = Worksheets("ObsBvul").Rows(i + 1).Columns(j)
            Next j
        Next i
    End If
    
    If RunFlags.InputCatch = True Then
        
        ReDim ObsCatch(StYear To EndYear, Nareas)
        'read Observed Catch
        For i = StYear To EndYear
            For area = 1 To Nareas
                ObsCatch(i, area) = Worksheets("ObsCatch").Rows(i - StYear + 1).Columns(area)
            Next area
        Next i
    End If
    
    If RunFlags.InputAbundance = True Then
        'read Observed Abundances
        i = 0
        Do
            i = i + 1
            j = Worksheets("ObsAbundance").Rows(i + 1).Columns(1)
        Loop While (j = Empty) = False
        If i = 1 Then MsgBox ("Include Abundance Data or change Input Abundance Flag")
        NObsAbundance = i - 1
        
        ReDim ObsAbundance(NObsAbundance, 4)
        
        For i = 1 To NObsAbundance
            For j = 1 To 4
                ObsAbundance(i, j) = Worksheets("ObsAbundance").Rows(i + 1).Columns(j)
            Next j
        Next i
    End If
End If

' Read Monte Carlo parameters
Nreplicates = Worksheets("Input").Rows(row_run_options + 2).Columns(2)

If RunFlags.Run_type = 1 And RecCV = 0 Then
    MsgBox ("Recruitment CV cannot be 0 when conditioning")
    Sheets("Input").Select
    End
End If

RecTimeCor = Worksheets("Input").Rows(row_Population_Dynamics + 3).Columns(2)

'RecSpaceCor = Worksheets("Input").Rows(row_Population_Dynamics + 4).Columns(2)
'posiblemente leer de otra worksheet

If RunFlags.ObsError_Survey Then SurveyCV = 0

If RunFlags.ProcError_Rec = 1 Then RecCV = 0

If RunFlags.ProcError_InitConditions = 1 Then InitialCV = 0


'Reading the Zvector in blocks of 50,000 elements, so far in two columns up to 100,000 as specified by the N_Zvector
h = 0
For j = 1 To 2
    For i = 1 To 50000
        h = h + 1
        Zvector(h) = Worksheets("ZZvector").Rows(i).Columns(j)
    Next i
Next j

End Sub
