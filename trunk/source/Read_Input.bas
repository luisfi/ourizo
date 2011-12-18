Attribute VB_Name = "Read_Input"
Option Explicit

Dim row_parameters As Integer, row_area_atributes As Integer, row_Population_Dynamics As Integer, _
    row_global_parameters As Integer, row_parameters_biol_region As Integer, row_parameters_area As Integer, _
    row_initial_conditions As Integer, row_connectivity As Integer, _
    row_management_control As Integer, row_run_options As Integer, row_catch_specification As Integer, row_effort_specification As Integer, _
    row_input_conditioning As Integer, row_reopening_conditions As Integer

Sub Read_Input()

row_run_options = 5
row_input_conditioning = 18
row_parameters = 30

    Nareas = Worksheets("Input").Rows(row_parameters + 1).Columns(2)
    StYear = Worksheets("Input").Rows(row_parameters + 2).Columns(2)
    EndYear = Worksheets("Input").Rows(row_parameters + 3).Columns(2)
    Nt = Worksheets("Input").Rows(row_parameters + 4).Columns(2)
    Stage = Worksheets("Input").Rows(row_parameters + 5).Columns(2)
    AgePlus = Worksheets("Input").Rows(row_parameters + 6).Columns(2)
    L1 = Worksheets("Input").Rows(row_parameters + 7).Columns(2)
    Linc = Worksheets("Input").Rows(row_parameters + 8).Columns(2)
    Nilens = Worksheets("Input").Rows(row_parameters + 9).Columns(2)
    Nyears = EndYear - StYear + 1

row_area_atributes = 41
row_Population_Dynamics = 46
row_global_parameters = 54
row_parameters_biol_region = 60
row_parameters_area = 72
row_initial_conditions = 80
row_management_control = 87
row_reopening_conditions = 4

row_connectivity = 135
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

    For Area = 1 To Nareas
        Surface(Area) = Worksheets("Input").Rows(row_area_atributes + 2).Columns(1 + Area)
        Lat(Area) = Worksheets("Input").Rows(row_area_atributes + 3).Columns(1 + Area)
        Lon(Area) = Worksheets("Input").Rows(row_area_atributes + 4).Columns(1 + Area)
    Next Area

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
    For Area = 1 To Nareas
        Bregion(Area) = Worksheets("Input").Rows(row_parameters_biol_region + 3).Columns(Area + 1)
      
        Linf(Area) = Worksheets("Input").Rows(row_parameters_biol_region + 4).Columns(Bregion(Area) + 1)
        k(Area) = Worksheets("Input").Rows(row_parameters_biol_region + 5).Columns(Bregion(Area) + 1)
        t0(Area) = Worksheets("Input").Rows(row_parameters_biol_region + 6).Columns(Bregion(Area) + 1)
        CVmu(Area) = Worksheets("Input").Rows(row_parameters_biol_region + 7).Columns(Bregion(Area) + 1)
        aW(Area) = Worksheets("Input").Rows(row_parameters_biol_region + 8).Columns(Bregion(Area) + 1)
        bW(Area) = Worksheets("Input").Rows(row_parameters_biol_region + 9).Columns(Bregion(Area) + 1)
        M(Area) = Worksheets("Input").Rows(row_parameters_biol_region + 10).Columns(Bregion(Area) + 1)
    Next Area

    For Area = 1 To Nareas
        Kcarga(Area) = Worksheets("Input").Rows(row_parameters_area + 1).Columns(Area + 1)
          Kcarga(Area) = Kcarga(Area) * Surface(Area)
        Rmax(Area) = Worksheets("Input").Rows(row_parameters_area + 2).Columns(Area + 1)
          Rmax(Area) = Rmax(Area) * Surface(Area)
        q(Area) = Worksheets("Input").Rows(row_parameters_area + 3).Columns(Area + 1)
        cost(Area) = Worksheets("Input").Rows(row_parameters_area + 4).Columns(Area + 1)
        If RunFlags.Growth_type = 1 Then
            gk(Area) = 1
        Else
            gk(Area) = Worksheets("Input").Rows(row_parameters_area + 5).Columns(Area + 1)
            Bthreshold(Area) = Worksheets("Input").Rows(row_parameters_area + 6).Columns(Area + 1)
            Bthreshold(Area) = Bthreshold(Area) * Kcarga(Area)
        End If
    
    Next Area

    RunFlags.Initial_Conditions = Worksheets("Input").Rows(row_initial_conditions + 1).Columns(2)
    For Area = 1 To Nareas
        HR_start(Area) = Worksheets("Input").Rows(row_initial_conditions + 2).Columns(Area + 1)
    Next Area

    For Area = 1 To Nareas
        For i = 1 To Nareas
            Connect(Area, i) = Worksheets("Input").Rows(row_connectivity + Area + 1).Columns(1 + i)
        Next i
    Next Area
              
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

For Area = 1 To Nareas
    Region(Area) = Worksheets("Input").Rows(row_management_control + 3).Columns(Area + 1)
    ClosedArea(StYear, Area) = Worksheets("Input").Rows(row_management_control + 4).Columns(1 + Area)
    Lfull(Area) = Worksheets("Input").Rows(row_management_control + 5).Columns(1 + Area)
    iLfull(Area) = ((Lfull(Area) - L1) / Linc) + 1
Next Area


RunFlags.Hstrategy = Worksheets("Input").Rows(row_management_control + 6).Columns(2)
TAC_TAE_HR = Worksheets("Input").Rows(row_management_control + 7).Columns(2)
MaxEffort = Worksheets("Input").Rows(row_management_control + 8).Columns(2)
Feedback = Worksheets("Input").Rows(row_management_control + 9).Columns(2)
TargetHR = Worksheets("Input").Rows(row_management_control + 10).Columns(2)


Nt_Season = Worksheets("Input").Rows(row_management_control + 11).Columns(2)
t_StSeason = Worksheets("Input").Rows(row_management_control + 12).Columns(2)
Nsurveys = Worksheets("Input").Rows(row_management_control + 29 + row_reopening_conditions).Columns(2)


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
        For Area = 1 To Nareas
            TAC_area(i, Area) = Worksheets("Input").Rows(row_catch_specification + 2 + (i - StYear)).Columns(2 + Nregions + Area)
        Next Area
    Next i
        
    PartialSurveyFlag = Worksheets("Input").Rows(row_management_control + 15).Columns(2)
    
    'if there is no feedback you can't have a partial survey
    If (Feedback = False And PartialSurveyFlag = True) Then
      PartialSurveyFlag = False
      MsgBox ("Inconsistent flags: setting PartialSurveyFlag to False")
    End If
    
    PulseHR = Worksheets("Input").Rows(row_management_control + 16).Columns(2)
    
    ReOpenConditionFlag = Worksheets("Input").Rows(row_management_control + 17).Columns(2)
    NOpenConditions = 4
    
    ReDim ReOpenCondition(NOpenConditions)
    ReDim ShortenTolerance(NOpenConditions)
    ReDim ReOpenConditionValues(NOpenConditions)
    
    For i = 1 To NOpenConditions
        ReOpenCondition(i) = Worksheets("Input").Rows(row_management_control + 17 + i).Columns(2)
        ShortenTolerance(i) = Worksheets("Input").Rows(row_management_control + 17 + i).Columns(3)
    Next i
    
    
    RestingTimeFlag = Worksheets("Input").Rows(row_management_control + 18 + row_reopening_conditions).Columns(2)
    
    
    If (RestingTimeFlag = True) Or (RunFlags.RotationType = 4) Then 'If areas to be ordered by restingtime or if rotation by period
        ReDim RestingTime(Nareas)
        ReDim RotationPeriod(Nareas)
        For Area = 1 To Nareas
            RestingTime(Area) = Worksheets("Input").Rows(row_management_control + 19 + row_reopening_conditions).Columns(1 + Area)
            RotationPeriod(Area) = Worksheets("Input").Rows(row_management_control + 20 + row_reopening_conditions).Columns(1 + Area)
        Next Area
    End If
    
    AdaptativeRotationFlag = Worksheets("Input").Rows(row_management_control + 21 + row_reopening_conditions).Columns(2)
    
Case 2  'By Area
    If (Nt_Season > 1) Then
        MsgBox ("Area_by_area strategy not implemented for Nt_Season> 1 ")
        End
    End If

    If Feedback = False Then
            'Public TAC_area y Effort
            If TAC_TAE_HR = 1 Then
                For i = StYear To EndYear
                    For Area = 1 To Nareas
                         TAC_area(i, Area) = Worksheets("Input").Rows(row_catch_specification + 2 + (i - StYear)).Columns(2 + Nregions + Area)
                    Next Area
                Next i
            
            ElseIf TAC_TAE_HR = 2 Then
                
                For i = StYear To EndYear
                    For Area = 1 To Nareas
                        TAE_area(i, Area) = Worksheets("Input").Rows(row_effort_specification + 2 + (i - StYear)).Columns(2 + Nregions + Area)
                    Next Area
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
    
    EffortDistributionFlag = Worksheets("Input").Rows(row_management_control + 23 + row_reopening_conditions).Columns(2)
        
    Select Case EffortDistributionFlag
    
    Case 1  'obsolete IFD
        Sens = Worksheets("Input").Rows(row_management_control + 24 + row_reopening_conditions).Columns(2)
        Ndias_beforeswitch = Worksheets("Input").Rows(row_management_control + 25 + row_reopening_conditions).Columns(2)
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

SurveyCV = Worksheets("Input").Rows(row_management_control + 33 + row_reopening_conditions).Columns(2)
InitialCV = Worksheets("Input").Rows(row_initial_conditions + 3).Columns(2)
RecCV = Worksheets("Input").Rows(row_Population_Dynamics + 2).Columns(2)

'Read Input Data for conditioning
If RunFlags.Run_type = 1 Then
    Nreplicates = 1
    
    If RunFlags.InputRec = True Then
        'read Observed Recruitment
        'q_rec = READ FROM INPUT
        
        For i = StYear To EndYear
            For Area = 1 To Nareas
                ObsRec(i, Area) = Worksheets("ObsRec").Rows((i - StYear + 1 + Nyears * (Area - 1))).Columns(1)
            Next Area
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
            For Area = 1 To Nareas
                ObsCatch(i, Area) = Worksheets("ObsCatch").Rows(i - StYear + 1).Columns(Area)
            Next Area
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
