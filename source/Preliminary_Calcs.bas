Attribute VB_Name = "Preliminary_Calcs"
Dim area As Integer, age As Integer, i As Integer, rr As Integer, i_area As Integer, Nopenareas As Integer
Dim mu_tmp As Double, sd_tmp As Double, pLTmp() As Double, year As Integer, IDopenarea() As Integer
Dim SBRXConectividad() As Double
Dim SBR0_avg As Double
Dim minSBRXConectividad As Double
Dim TotalSurface As Double, Wvul As Double, R0total As Double
Dim yr As Integer, t As Integer, MaxNareas_Region As Integer

'#################################################################################
'#                                                              INITIALIZE                                                                                            #
'#################################################################################

Sub Initialize_variables()
Attribute Initialize_variables.VB_ProcData.VB_Invoke_Func = " \n14"
    SimEndYear = StYear + 200
  '  Nyears = EndYear - StYear + 1
    Nages = AgePlus - Stage + 1 '
    NpulsosMax = Nt * Nages

ReDim Btotal(StYear - Stage + 1 To SimEndYear + 1, Nareas)
ReDim Bmature(StYear - Stage To SimEndYear + 1, Nareas)
ReDim SurveyBtot(Nsurveys, StYear - Stage + 1 To SimEndYear + 1, Nareas)
ReDim SurveyBvul(Nsurveys, StYear - Stage + 1 To SimEndYear + 1, Nareas)
ReDim SurveyMat(Nsurveys, StYear - Stage + 1 To SimEndYear + 1, Nareas)
ReDim SurveyNage(Nsurveys, StYear - Stage + 1 To SimEndYear + 1, Nareas, Nages)
ReDim SurveyNtot(Nsurveys, StYear - Stage + 1 To SimEndYear + 1, Nareas)
ReDim SurveypL(Nsurveys, StYear - Stage + 1 To SimEndYear + 1, Nareas, Nilens)
ReDim Bvulnerable(StYear - Stage + 1 To SimEndYear + 1, Nareas)
ReDim N(StYear To SimEndYear + 1, Nareas, Stage To AgePlus)
ReDim mu(StYear To SimEndYear + 1, Nareas, Stage To AgePlus)
ReDim sd(StYear To SimEndYear + 1, Nareas, Stage To AgePlus)
ReDim w(StYear To SimEndYear + 1, Nareas, Stage To AgePlus)
ReDim Larvae(StYear - Stage To SimEndYear, Nareas)
ReDim Settlers(StYear - Stage + 1 To SimEndYear + Stage, Nareas)
ReDim pL(StYear To SimEndYear, Nareas, Nilens)
ReDim Rdev(StYear To EndYear, Nareas)

ReDim g(Nareas)
ReDim Bg0(Nareas)

'JLV July 20, 2007

ReDim TAE_area(StYear To EndYear, Nareas)
' ReDim TAC_area(StYear To EndYear, Nareas)

ReDim Nfracs(Nareas, Stage To AgePlus)
ReDim Catch(StYear To EndYear, Nareas)
ReDim effort(StYear To EndYear, Nareas)

ReDim OpenMonth(StYear To EndYear, Nt)
ReDim Nareas_region(Nregions)
ReDim AnnualCatch(Nregions)
ReDim ClosedRegion(Nregions)
ReDim ClosedRegionTmp(Nregions)
ReDim NTmp(Nareas, Stage To AgePlus)
ReDim BvulTmp(Nareas)
ReDim BtotTmp(Nareas)
ReDim muTmp(Nareas, Stage To AgePlus)
ReDim sdTmp(Nareas, Stage To AgePlus)
ReDim WTmp(Nareas, Stage To AgePlus)
ReDim flag_Partial_Rec(Nareas, Stage To AgePlus)
ReDim Z(Nareas, Stage To AgePlus, NpulsosMax)
ReDim frac(Nareas, Stage To AgePlus, NpulsosMax)
ReDim pLage(Nareas, Stage To AgePlus, Nilens)
ReDim pLageplus(Nareas, Nilens)
ReDim pLStAge(Nareas, Nilens)
ReDim W_L(Nareas, Nilens)
ReDim l(Nilens)
ReDim Alpha0(Nareas)
ReDim Alpha(Nareas)
ReDim Beta0(Nareas)
ReDim Beta(Nareas)
ReDim Rho(Nareas)
ReDim pLTmp(Nilens)
ReDim WvulStage(Nareas)
ReDim FracSel(Nareas, Stage To AgePlus)
ReDim FracSelStAge(Nareas)
ReDim HRTmp(Nareas)
ReDim Kcarga_adults(Nareas)
ReDim Atlas(Nareas)


For area = 1 To Nareas
    For year = StYear To EndYear
         Catch(year, area) = 0
    Next year
        
    For age = Stage To AgePlus
            Nfracs(area, age) = 0
    Next age

       'Optar por tipo de crecimiento
        Select Case RunFlags.Growth_type
          Case 1
            'Density-independent growth
               
            gk(area) = 1
                        
          Case 2
            
            'Lineal density dependence
         
            Bg0(area) = (Kcarga(area) - Bthreshold(area)) / (1 - gk(area)) + Bthreshold(area)
                             
        End Select
            Rho(area) = Exp(-k(area))
            Alpha(area) = (1 - Rho(area)) * Linf(area) * gk(area)
            Beta(area) = 1 - (1 - Rho(area)) * gk(area)
                             
            Alpha0(area) = Alpha(area)
            Beta0(area) = Beta(area)
        
        'Estimate mu@age StAge and W@age StAge per area
        mu(StYear, area, Stage) = Linf(area) * (1 - Exp(-k(area) * (Stage - t0(area))))
        sd(StYear, area, Stage) = CVmu(area) * mu(StYear, area, Stage)
        muTmp(area, Stage) = mu(StYear, area, Stage)
        sdTmp(area, Stage) = sd(StYear, area, Stage)
        
        For ilen = 1 To Nilens
            l(ilen) = L1 + (Linc * (ilen - 1))
            W_L(area, ilen) = aW(area) * l(ilen) ^ bW(area)
        Next ilen
        
Next area

For area = 1 To Nareas
    For age = Stage To AgePlus
        flag_Partial_Rec(area, age) = 1
    Next age
Next area

For area = 1 To Nareas
    For age = Stage To AgePlus
        For i = 1 To NpulsosMax
            Z(area, age, i) = 0
            frac(area, age, i) = 1
        Next i
    Next age
Next area


For yr = StYear To EndYear
    For t = t_StSeason To (t_StSeason + Nt_Season - 1)
        OpenMonth(yr, t) = True
    Next t
    For area = 1 To Nareas
        'at this point this is just to close permanent areas
        ClosedArea(yr, area) = ClosedArea(StYear, area)
    Next area
Next yr

For rr = 1 To Nregions
   ClosedRegion(rr) = True
Next rr
        
i_area = 0
For area = 1 To Nareas
    If ClosedArea(StYear, area) = False Then
        ClosedRegion(Region(area)) = False
        i_area = i_area + 1
        ReDim Preserve IDopenarea(i_area)
        IDopenarea(i_area) = area
    End If
Next area
Nopenareas = i_area

For rr = 1 To Nregions
   ClosedRegionTmp(rr) = ClosedRegion(rr)
Next rr


If RunFlags.Hstrategy = 1 Then ' If rotational strategy
    'copy the open areas and resting times in sheet Calcs
            For i_area = 1 To Nopenareas
                Worksheets("Calcs").Rows(i_area).Columns(1) = IDopenarea(i_area)
                Worksheets("Calcs").Rows(i_area).Columns(2) = RestingTime(IDopenarea(i_area))
            Next i_area
     
    'Sort the open areas by resting time
            Worksheets("Calcs").Activate
            Range(Cells(1, 1), Cells(Nopenareas, 2)).Select
            Selection.Sort Key1:=Range("B1"), Order1:=xlDescending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
    
            For i_area = 1 To Nopenareas
                IDopenarea(i_area) = Worksheets("Calcs").Rows(i_area).Columns(1)
            Next i_area
End If
 
MaxNareas_Region = 1

For i_area = 1 To Nopenareas

    rr = Region(IDopenarea(i_area))
    
        Nareas_region(rr) = Nareas_region(rr) + 1
        
        If Nareas_region(rr) > MaxNareas_Region Then MaxNareas_Region = Nareas_region(rr)
        
        ReDim Preserve Candidate_areas(Nregions, MaxNareas_Region)
        
        Candidate_areas(rr, Nareas_region(rr)) = IDopenarea(i_area)
    
Next i_area

For rr = 1 To Nregions
    AnnualCatch(Nregions) = 0
    For i = 1 To Nareas_region(rr)
            Worksheets("Calcs").Rows(rr + 10).Columns(i + 10) = Candidate_areas(rr, i)
    Next i
Next rr


'IF ROTATION THEN CALCULATE AREA
TotalSurface = 0
For area = 1 To Nareas
    TotalSurface = TotalSurface + Surface(area)
Next area

TargetSurface = TargetHR * TotalSurface
PulseHRadjust = 1

End Sub

'#################################################################################
'#                             CARRYING CAPACITY                                                                              #
'#################################################################################

Sub Set_Carrying_Capacity()
Attribute Set_Carrying_Capacity.VB_ProcData.VB_Invoke_Func = " \n14"
  
  'Calculates all variables at carrying capacity starting from Kcarga(Area) calculated at Read_Input
  Dim pLageplusTmp()

  Dim Amax() As Double, Amaxtemp As Integer, muStyeartemp() As Double, NStyeartemp() As Double
  ReDim pLageplusTmp(Nilens)


'ReDim Settlers(StYear - StAge + 1 To EndYear + StAge, Nareas)
  ReDim Amax(Nareas)
  ReDim R0(Nareas), SBR0(Nareas), SB0(StYear To StYear, Nareas), BR0(Nareas), VB0(Nareas)
  ReDim SBRXConectividad(Nareas)
  
'Aca estimo cuantos anios proyectar largos y pesos para representar adecuadamente los valores inciales del plusgroup de manera de
'proyectar hasta una edad Amax correspondiente al 0.001 de la abundandia inicial al reclutamiento
  For area = 1 To Nareas
            
''Amax(Area) = -Log(0.0001) / M(Area)
     Amax(area) = 1000
     Amaxtemp = Amax(area)
   
     ReDim muStyeartemp(Nareas, Stage To Amaxtemp)
     ReDim NStyeartemp(Nareas, Stage To Amaxtemp)
           
     muStyeartemp(area, Stage) = mu(StYear, area, Stage)
     NStyeartemp(area, Stage) = 1
                       
     For age = Stage To AgePlus
          muStyeartemp(area, age + 1) = Alpha(area) + Beta(area) * muStyeartemp(area, age)
          NStyeartemp(area, age + 1) = NStyeartemp(area, age) * Exp(-M(area))
          mu(StYear, area, age) = muStyeartemp(area, age)
          N(StYear, area, age) = NStyeartemp(area, age)
          muTmp(area, age) = mu(StYear, area, age)
          sdTmp(area, age) = CVmu(area) * mu(StYear, area, age)
          
          Call M8_Library.Norm(area, age)
            
          FracSel(area, age) = 0
          For ilen = iLfull(area) To Nilens
                FracSel(area, age) = FracSel(area, age) + pLage(area, age, ilen)
          Next ilen
        
        'Debug.Print FracSel(area,age)
     Next age
                       
    'Initialize calculation of pL of  Plus group and store pL for StAge
             
     For ilen = 1 To Nilens
        pLageplus(area, ilen) = pLage(area, AgePlus, ilen) * N(StYear, area, AgePlus)
        pLStAge(area, ilen) = pLage(area, Stage, ilen)
     Next ilen
    
     FracSelStAge(area) = FracSel(area, Stage)
                
     Select Case RunFlags.VirginAgePlus
    
     Case 1
    ''''AGE-BASED FORMULATION FOR PLUS GROUP
    'Compute size composition and average of mu of the Plus group
    
       mu(StYear, area, AgePlus) = mu(StYear, area, AgePlus) * N(StYear, area, AgePlus)
            
       For age = AgePlus + 1 To Amaxtemp
                
     'Calcular pLtmp only for this plus group
          muStyeartemp(area, age) = Alpha(area) + Beta(area) * muStyeartemp(area, age - 1)
          NStyeartemp(area, age) = NStyeartemp(area, age - 1) * Exp(-M(area))
          mu_tmp = muStyeartemp(area, age)
          sd_tmp = CVmu(area) * mu_tmp
          intfact = 0
          For ilen = 1 To Nilens
              pLTmp(ilen) = Exp(-0.5 * ((l(ilen) - mu_tmp) / sd_tmp) ^ 2)
              intfact = intfact + pLTmp(ilen)
          Next
    
          For ilen = 1 To Nilens
             pLageplus(area, ilen) = pLageplus(area, ilen) + pLTmp(ilen) * NStyeartemp(area, age) / intfact
          Next ilen
                
          mu(StYear, area, AgePlus) = mu(StYear, area, AgePlus) + muStyeartemp(area, age) * NStyeartemp(area, age)
          N(StYear, area, AgePlus) = N(StYear, area, AgePlus) + NStyeartemp(area, age)
                
       Next age
    
     Case 2
    ''''LENGTH-BASED FORMULATION to calculate size composition and mu of PLUS GROUP
    '
      Dim Lplus As Double
      Dim ilenplus As Integer
      
      
      For age = AgePlus + 1 To Amaxtemp
          
          For ilen = 1 To Nilens
                   pLageplusTmp(ilen) = 0
          Next ilen
          
          'now growth
           ilen = 1

           While l(ilen) < Linf(area) And ilen < Nilens
   '           Lplus = Alpha(Area) + Beta(Area) * (l(ilen) + 0.5 * Linc)
              Lplus = Alpha(area) + Beta(area) * l(ilen) 'fixed with Ines- l(ilen) is the center of interval
              ilenplus = 1 + (Lplus - L1) / Linc     'this rounds the number (doesn't truncate)
              If (ilenplus > Nilens) Then ilenplus = Nilens
              pLageplusTmp(ilenplus) = pLageplusTmp(ilenplus) + pLageplus(area, ilen)
              ilen = ilen + 1
           Wend
            
           For i = ilen To Nilens       'NB! these are for l(ilen) >= Linf
                pLageplusTmp(i) = pLageplusTmp(i) + pLageplus(area, i)
           Next i
            
           For ilen = 1 To Nilens
                pLageplus(area, ilen) = pLageplusTmp(ilen) * Exp(-M(area)) + pLage(area, AgePlus, ilen) * NStyeartemp(area, AgePlus)
                ' Debug.Print pLageplus(area, ilen)
           Next ilen
           
           NStyeartemp(area, age) = NStyeartemp(area, age - 1) * Exp(-M(area))
           N(StYear, area, AgePlus) = N(StYear, area, AgePlus) + NStyeartemp(area, age)
      Next age
              
      For ilen = 1 To Nilens
           Debug.Print pLageplus(area, ilen)
           mu(StYear, area, AgePlus) = mu(StYear, area, AgePlus) + l(ilen) * pLageplus(area, ilen)
      Next ilen
                                   
     End Select
 '''''''''''''''''''' END AGE vs SIZE FORMULATION for size comp of Plus group
 
     mu(StYear, area, AgePlus) = mu(StYear, area, AgePlus) / N(StYear, area, AgePlus)
    
     For ilen = 1 To Nilens
        pLageplus(area, ilen) = pLageplus(area, ilen) / N(StYear, area, AgePlus)
        pLage(area, AgePlus, ilen) = pLageplus(area, ilen)
     Next ilen
     
     'Get FracMat() vector from Maturity Function
     Call M5_Popdyn.Maturity(AgeFullMature, FracMat)
                
     BR0(area) = 0
     SBR0(area) = 0
    
     For age = Stage To AgePlus
        w(StYear, area, age) = 0
            
        For ilen = 1 To Nilens
           w(StYear, area, age) = w(StYear, area, age) + W_L(area, ilen) * pLage(area, age, ilen)
        Next ilen
            
        BR0(area) = BR0(area) + w(StYear, area, age) * N(StYear, area, age)  'total biomass per recruit- N is here a survival fraction
        SBR0(area) = SBR0(area) + w(StYear, area, age) * N(StYear, area, age) * FracMat(age)  'spawning biomass per recruit

     Next age
        
     R0(area) = Kcarga(area) / BR0(area)
     N(StYear, area, Stage) = R0(area)
     Bmature(StYear, area) = R0(area) * SBR0(area)
                
     Btotal(StYear, area) = 0
     Bvulnerable(StYear, area) = 0
          
     For age = Stage + 1 To AgePlus
         Wvul = 0
         For ilen = iLfull(area) To Nilens
             Wvul = Wvul + W_L(area, ilen) * pLage(area, age, ilen)
         Next ilen
            
         N(StYear, area, age) = R0(area) * N(StYear, area, age)
    'NB: these are missing contribution of Stage
         Btotal(StYear, area) = Btotal(StYear, area) + N(StYear, area, age) * w(StYear, area, age)
         Bvulnerable(StYear, area) = Bvulnerable(StYear, area) + N(StYear, area, age) * Wvul * FracSel(area, age)
   
              'Debug.Print Bvulnerable(StYear, area)
              'Debug.Print Wvul
     Next age   'this loop is from StAge+1
                
     Wvul = 0
     For ilen = iLfull(area) To Nilens
         Wvul = Wvul + W_L(area, ilen) * pLage(area, Stage, ilen)
     Next ilen
     WvulStage(area) = Wvul

     VB0(area) = Bvulnerable(StYear, area) + R0(area) * WvulStage(area) * FracSel(area, Stage)
     SB0(StYear, area) = Bmature(StYear, area)   'esta inclute Stage
    
  Next area

  SBR0_avg = 0
  R0total = 0
  VB0_all = 0
  SB0_all = 0
    
  For area = 1 To Nareas
       VB0_all = VB0_all + VB0(area)
       SB0_all = SB0_all + SB0(StYear, area)
       SBR0_avg = SBR0_avg + R0(area) * SBR0(area)
       R0total = R0total + R0(area)
  Next area
    
    SBR0_avg = SBR0_avg / R0total
        
   For area = 1 To Nareas
     SBRXConectividad(area) = 0
     For i = 1 To Nareas
         'SBRXConectividad(Area) = SBRXConectividad(Area) + Connect(Area, i) * SBR0(i)
         SBRXConectividad(area) = SBRXConectividad(area) + Connect(area, i) * SBR0(i) * R0(i) / R0(area)
     Next i

       ' Debug.Print SBRXConectividad(Area)
   Next area
    
  'compute minimum
   minSBRXConectividad = SBRXConectividad(1)

     For area = 2 To Nareas
          If SBRXConectividad(area) < minSBRXConectividad Then minSBRXConectividad = SBRXConectividad(area)
     Next area
    
   ProdXB = (1 / minSBRXConectividad) * Lambda_ProdXB
    
   Call M6_Prod_Alloc_Larvae.Prod_Alloc_Larvae(StYear, SB0) 'in StYear B will be at carrying capacity

   For area = 1 To Nareas
        
      For year = StYear To StYear + Stage - 1
            Settlers(year, area) = Settlers(StYear + Stage, area)
         '      Debug.Print Settlers(year, area)
      Next year
   Next area

  
  Call Input_Output.Print_Initial_Conditions("Carrying_Capacity")

End Sub
'#######################################################################################
'#                            VIRGIN CONDITIONS
'#  Calculate virgin conditions (unharvested equilibrium) by projecting population
'#  forward from carrying capacity. Note that if Lambda < 1 some areas will be limited by
'#  larval supply and will be below K.
'#######################################################################################

Sub Set_Virgin_Conditions()

Dim simyear As Integer
ReDim SBvirgin(Nareas), VBvirgin(Nareas)  'there are different from VB0 and SB0

' Biomass in StYear is at carrying capacity calculated in Set_Carrying_Capacity

' Simulate with harvest rate 0
        SimEndYear = StYear + 200
    
    'Start from K and simulate forward with no harvest
    For area = 1 To Nareas
          HRTmp(area) = 0
    Next area
                    
    Call Preliminary_Calcs.Initialize_tmp_variables  'these are the state variables used in the pop dyn loops
    
    For simyear = StYear To SimEndYear - 1
               
        Call M4_Calc_Recruits.Deterministic_Recruits(simyear)
            
        Call M6_Prod_Alloc_Larvae.Prod_Alloc_Larvae(simyear)
          
        Call M5_Popdyn.PopDyn(simyear)
                        
        Call M2_AnnualUpdate.Annual_update(simyear)
    
     Next simyear
     
'initialize variables for initial condition simulation
     For area = 1 To Nareas
         For age = Stage To AgePlus
              N(StYear, area, age) = N(SimEndYear, area, age)
              mu(StYear, area, age) = mu(SimEndYear, area, age)
              sd(StYear, area, age) = sd(SimEndYear, area, age)
              w(StYear, area, age) = w(SimEndYear, area, age)
          Next age
          
          N(StYear, area, Stage) = N(SimEndYear - 1, area, Stage)
          Bmature(StYear, area) = Bmature(SimEndYear - 1, area)
          'NB next biomasses are missing Stage which is OK - will be added at start of pop dyn loop
          Btotal(StYear, area) = Btotal(SimEndYear, area)
          Bvulnerable(StYear, area) = Bvulnerable(SimEndYear, area)
      
          
          For year = StYear To StYear + Stage - 1
              Settlers(year, area) = Settlers(SimEndYear - 1, area)
          Next year
          
          VBvirgin(area) = Bvulnerable(StYear, area) + N(StYear, area, Stage) * WvulStage(area) * FracSel(area, Stage)
          SBvirgin(area) = Bmature(StYear, area)  'includes Stage because it was calculated within Prod_Alloc_Larvae
          
     Next area
                        
     Call Input_Output.Print_Initial_Conditions("Virgin_Conditions")


VBvirgin_all = 0
SBvirgin_all = 0
For area = 1 To Nareas
    VBvirgin_all = VBvirgin_all + VBvirgin(area)
    SBvirgin_all = SBvirgin_all + SBvirgin(area)
Next area

End Sub



'#################################################################################
'#                         INITIAL CONDITIONS                                                                         #
'#################################################################################

Sub Set_InitialConditions()
Attribute Set_InitialConditions.VB_ProcData.VB_Invoke_Func = " \n14"
Dim year As Integer, area As Integer, simyear As Integer

Select Case RunFlags.Initial_Conditions

  Case 1
     'start all areas at carrying capacity
     Call Input_Output.Read_Initial_Conditions("Carrying_Capacity")
     Call Input_Output.Print_Initial_Conditions("Initial_Conditions")
    
  Case 2
     'start at HR=0 equilibrium conditions
     Call Input_Output.Read_Initial_Conditions("Virgin_Conditions")
     Call Input_Output.Print_Initial_Conditions("Initial_Conditions")
        
  Case 3
     'start at equilibroum conditions for HR_start solved here by simulations
        SimEndYear = StYear + 200
    
    'Start from Virgin and simulate forward under constant harvest rate = HR_start(area)
  
    For area = 1 To Nareas
          HRTmp(area) = HR_start(area)
    Next area
                
    For simyear = StYear To SimEndYear - 1
               
        Call M4_Calc_Recruits.Deterministic_Recruits(simyear)
            
        Call M6_Prod_Alloc_Larvae.Prod_Alloc_Larvae(simyear)
          
        Call M5_Popdyn.PopDyn(simyear)
                        
        Call M2_AnnualUpdate.Annual_update(simyear)
    
     Next simyear
 
     For area = 1 To Nareas
         For age = Stage To AgePlus
              N(StYear, area, age) = N(SimEndYear, area, age)
              mu(StYear, area, age) = mu(SimEndYear, area, age)
              sd(StYear, area, age) = sd(SimEndYear, area, age)
              w(StYear, area, age) = w(SimEndYear, area, age)
          Next age
          
          N(StYear, area, Stage) = N(SimEndYear - 1, area, Stage) 'NB use SimEndYear-1 because after annual update N at Stage is zero
          Bmature(StYear, area) = Bmature(SimEndYear - 1, area)
          Btotal(StYear, area) = Btotal(SimEndYear, area)
          Bvulnerable(StYear, area) = Bvulnerable(SimEndYear, area)

                               
          For year = StYear To StYear + Stage
              Settlers(year, area) = Settlers(SimEndYear - 1, area)
          Next year
     Next area
                        
     Call Input_Output.Print_Initial_Conditions("Initial_Conditions")
  
  Case 4
        'Read from arbitrary input file (e.g. estimated from conditioning)
       
     Call Input_Output.Read_Initial_Conditions("Initial_Conditions")
  
  End Select

End Sub
Sub Rescale_parameters()
Attribute Rescale_parameters.VB_ProcData.VB_Invoke_Func = " \n14"
For area = 1 To Nareas
        M(area) = M(area) / Nt
        k(area) = k(area) / Nt
        Rho(area) = Exp(-k(area))
Next area
End Sub

'#################################################################################
'#                INITIALIZE TEMP VARIABLES BEFORE SIMULATION LOOPS                                                                          #
'#################################################################################

Sub Initialize_tmp_variables()

  For area = 1 To Nareas
  '     BvulTmp(Area) = Bvulnerable(StYear, Area)
  '     BtotTmp(Area) = Btotal(StYear, Area)
       
       For age = Stage To AgePlus
           NTmp(area, age) = N(StYear, area, age)
           muTmp(area, age) = mu(StYear, area, age)
           sdTmp(area, age) = CVmu(area) * muTmp(area, age)
           WTmp(area, age) = w(StYear, area, age)
       Next age
             
  Next area

End Sub
