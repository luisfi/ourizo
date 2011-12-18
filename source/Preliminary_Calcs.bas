Attribute VB_Name = "Preliminary_Calcs"
Dim Area As Integer, age As Integer, i As Integer, rr As Integer, i_area As Integer, Nopenareas As Integer
Dim mu_tmp As Double, sd_tmp As Double, pLTmp() As Double
Dim Amax() As Double, Amaxtemp As Integer, muStyeartemp() As Double, _
NStyeartemp() As Double, year As Integer, IDopenarea() As Integer
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
ReDim Survey(Nsurveys, StYear - Stage + 1 To SimEndYear + 1, Nareas)
ReDim SurveyAll(StYear - Stage + 1 To SimEndYear + 1)
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
ReDim Flag_Rec_Fish(Nareas, Stage To AgePlus)
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


For Area = 1 To Nareas
    For year = StYear To EndYear
        Bmature(year, Area) = 0
        Catch(year, Area) = 0
    Next year
        
        For age = Stage To AgePlus
            Nfracs(Area, age) = 0
        Next age

       'Optar por tipo de crecimiento
        Select Case RunFlags.Growth_type
          Case 1
            'Density-independent growth
               
            gk(Area) = 1
                        
          Case 2
            
            'Lineal density dependence
         
            Bg0(Area) = (Kcarga(Area) - Bthreshold(Area)) / (1 - gk(Area)) + Bthreshold(Area)
                             
        End Select
            Rho(Area) = Exp(-k(Area))
            Alpha(Area) = (1 - Rho(Area)) * Linf(Area) * gk(Area)
            Beta(Area) = 1 - (1 - Rho(Area)) * gk(Area)
                             
            Alpha0(Area) = Alpha(Area)
            Beta0(Area) = Beta(Area)
        
        'Estimate mu@age StAge and W@age StAge per area
        mu(StYear, Area, Stage) = Linf(Area) * (1 - Exp(-k(Area) * (Stage - t0(Area))))
        sd(StYear, Area, Stage) = CVmu(Area) * mu(StYear, Area, Stage)
        muTmp(Area, Stage) = mu(StYear, Area, Stage)
        sdTmp(Area, Stage) = sd(StYear, Area, Stage)
        
        For ilen = 1 To Nilens
            l(ilen) = L1 + (Linc * (ilen - 1))
            W_L(Area, ilen) = aW(Area) * l(ilen) ^ bW(Area)
        Next ilen
        
Next Area

For Area = 1 To Nareas
    For age = Stage To AgePlus
        Flag_Rec_Fish(Area, age) = 1
    Next age
Next Area

For Area = 1 To Nareas
    For age = Stage To AgePlus
        For i = 1 To NpulsosMax
            Z(Area, age, i) = 0
            frac(Area, age, i) = 1
        Next i
    Next age
Next Area


For yr = StYear To EndYear
    For t = t_StSeason To (t_StSeason + Nt_Season - 1)
        OpenMonth(yr, t) = True
    Next t
    For Area = 1 To Nareas
        'at this point this is just to close permanent areas
        ClosedArea(yr, Area) = ClosedArea(StYear, Area)
    Next Area
Next yr

For rr = 1 To Nregions
   ClosedRegion(rr) = True
Next rr
        
i_area = 0
For Area = 1 To Nareas
    If ClosedArea(StYear, Area) = False Then
        ClosedRegion(Region(Area)) = False
        i_area = i_area + 1
        ReDim Preserve IDopenarea(i_area)
        IDopenarea(i_area) = Area
    End If
Next Area
Nopenareas = i_area

For rr = 1 To Nregions
   ClosedRegionTmp(rr) = ClosedRegion(rr)
Next rr


If RestingTimeFlag = True Then
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
For Area = 1 To Nareas
    TotalSurface = TotalSurface + Surface(Area)
Next Area

TargetSurface = TargetHR * TotalSurface
PulseHRadjust = 1

End Sub

'#################################################################################
'#                             CARRYING CAPACITY                                                                              #
'#################################################################################

Sub Set_Carrying_Capacity()
Attribute Set_Carrying_Capacity.VB_ProcData.VB_Invoke_Func = " \n14"

Dim pLageplusTmp()
Dim SB0() As Double
ReDim pLageplusTmp(Nilens)


'ReDim Settlers(StYear - StAge + 1 To EndYear + StAge, Nareas)
ReDim Amax(Nareas)
ReDim R0(Nareas), SBR0(Nareas), SB0(StYear To StYear, Nareas), BR0(Nareas), VB0(Nareas)

ReDim SBRXConectividad(Nareas)

'Aca estimo cuantos anios proyectar largos y pesos para representar adecuadamente los valores inciales del plusgroup de manera de
'proyectar hasta una edad Amax correspondiente al 0.001 de la abundandia inicial al reclutamiento
For Area = 1 To Nareas
            
''Amax(Area) = -Log(0.0001) / M(Area)
Amax(Area) = 1000
Amaxtemp = Amax(Area)
    
ReDim muStyeartemp(Nareas, Stage To Amaxtemp)
ReDim NStyeartemp(Nareas, Stage To Amaxtemp)
            
  muStyeartemp(Area, Stage) = mu(StYear, Area, Stage)
  NStyeartemp(Area, Stage) = 1
                       
   
    For age = Stage To AgePlus
          muStyeartemp(Area, age + 1) = Alpha(Area) + Beta(Area) * muStyeartemp(Area, age)
          NStyeartemp(Area, age + 1) = NStyeartemp(Area, age) * Exp(-M(Area))
          mu(StYear, Area, age) = muStyeartemp(Area, age)
          N(StYear, Area, age) = NStyeartemp(Area, age)
          muTmp(Area, age) = mu(StYear, Area, age)
          sdTmp(Area, age) = CVmu(Area) * mu(StYear, Area, age)
          
          Call M8_Library.Norm(Area, age)
            
          FracSel(Area, age) = 0
          For ilen = iLfull(Area) To Nilens
                FracSel(Area, age) = FracSel(Area, age) + pLage(Area, age, ilen)
          Next ilen
        
        'Debug.Print FracSel(area,age)
    Next age
                       
    'Initialize calculation of pL of  Plus group and store pL for StAge
             
    For ilen = 1 To Nilens
       pLageplus(Area, ilen) = pLage(Area, AgePlus, ilen) * N(StYear, Area, AgePlus)
       pLStAge(Area, ilen) = pLage(Area, Stage, ilen)
    Next ilen
    
    FracSelStAge(Area) = FracSel(Area, Stage)
                
    Select Case RunFlags.VirginAgePlus
    
    Case 1
    ''''AGE-BASED FORMULATION FOR PLUS GROUP
    'Compute size composition and average of mu of the Plus group
    
    mu(StYear, Area, AgePlus) = mu(StYear, Area, AgePlus) * N(StYear, Area, AgePlus)
            
    For age = AgePlus + 1 To Amaxtemp
                
     'Calcular pLtmp only for this plus group
          muStyeartemp(Area, age) = Alpha(Area) + Beta(Area) * muStyeartemp(Area, age - 1)
          NStyeartemp(Area, age) = NStyeartemp(Area, age - 1) * Exp(-M(Area))
          mu_tmp = muStyeartemp(Area, age)
          sd_tmp = CVmu(Area) * mu_tmp
          intfact = 0
          For ilen = 1 To Nilens
              pLTmp(ilen) = Exp(-0.5 * ((l(ilen) - mu_tmp) / sd_tmp) ^ 2)
              intfact = intfact + pLTmp(ilen)
          Next
    
          For ilen = 1 To Nilens
             pLageplus(Area, ilen) = pLageplus(Area, ilen) + pLTmp(ilen) * NStyeartemp(Area, age) / intfact
          Next ilen
                
          mu(StYear, Area, AgePlus) = mu(StYear, Area, AgePlus) + muStyeartemp(Area, age) * NStyeartemp(Area, age)
          N(StYear, Area, AgePlus) = N(StYear, Area, AgePlus) + NStyeartemp(Area, age)
                
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

           While l(ilen) < Linf(Area) And ilen < Nilens
   '           Lplus = Alpha(Area) + Beta(Area) * (l(ilen) + 0.5 * Linc)
              Lplus = Alpha(Area) + Beta(Area) * l(ilen) 'fixed with Ines- l(ilen) is the center of interval
              ilenplus = 1 + (Lplus - L1) / Linc     'this rounds the number (doesn't truncate)
              If (ilenplus > Nilens) Then ilenplus = Nilens
              pLageplusTmp(ilenplus) = pLageplusTmp(ilenplus) + pLageplus(Area, ilen)
              ilen = ilen + 1
           Wend
            
           For i = ilen To Nilens       'NB! these are for l(ilen) >= Linf
                pLageplusTmp(i) = pLageplusTmp(i) + pLageplus(Area, i)
           Next i
            
           For ilen = 1 To Nilens
                pLageplus(Area, ilen) = pLageplusTmp(ilen) * Exp(-M(Area)) + pLage(Area, AgePlus, ilen) * NStyeartemp(Area, AgePlus)
                ' Debug.Print pLageplus(area, ilen)
           Next ilen
           
           NStyeartemp(Area, age) = NStyeartemp(Area, age - 1) * Exp(-M(Area))
           N(StYear, Area, AgePlus) = N(StYear, Area, AgePlus) + NStyeartemp(Area, age)
      Next age
              
      For ilen = 1 To Nilens
           Debug.Print pLageplus(Area, ilen)
           mu(StYear, Area, AgePlus) = mu(StYear, Area, AgePlus) + l(ilen) * pLageplus(Area, ilen)
      Next ilen
                                   
    End Select
 '''''''''''''''''''' END AGE vs SIZE FORMULATION for size comp of Plus group
 
    mu(StYear, Area, AgePlus) = mu(StYear, Area, AgePlus) / N(StYear, Area, AgePlus)
    
    For ilen = 1 To Nilens
        pLageplus(Area, ilen) = pLageplus(Area, ilen) / N(StYear, Area, AgePlus)
        pLage(Area, AgePlus, ilen) = pLageplus(Area, ilen)
                    
'Debug.Print pLage(area, AgePlus, ilen)
                    
    Next ilen
                
    BR0(Area) = 0
        
    For age = Stage To AgePlus
        w(StYear, Area, age) = 0
            
        For ilen = 1 To Nilens
           w(StYear, Area, age) = w(StYear, Area, age) + W_L(Area, ilen) * pLage(Area, age, ilen)
        Next ilen
            
        BR0(Area) = BR0(Area) + w(StYear, Area, age) * N(StYear, Area, age)
     Next age
        
     R0(Area) = Kcarga(Area) / BR0(Area)
     N(StYear, Area, Stage) = R0(Area)
                                                                               
     'Get FracMat() vector from Maturity Function
     Call M5_Popdyn.Maturity(AgeFullMature, FracMat)
                
        Btotal(StYear, Area) = 0
        Bvulnerable(StYear, Area) = 0
        
        For age = Stage + 1 To AgePlus
            Wvul = 0
            For ilen = iLfull(Area) To Nilens
                Wvul = Wvul + W_L(Area, ilen) * pLage(Area, age, ilen)
''Debug.Print pLage(Area, age, ilen)
            Next ilen
            
            N(StYear, Area, age) = R0(Area) * N(StYear, Area, age)
            Btotal(StYear, Area) = Btotal(StYear, Area) + N(StYear, Area, age) * w(StYear, Area, age)
            Bvulnerable(StYear, Area) = Bvulnerable(StYear, Area) + N(StYear, Area, age) * Wvul * FracSel(Area, age)
            Bmature(StYear, Area) = Bmature(StYear, Area) + N(StYear, Area, age) * w(StYear, Area, age) * FracMat(age)
                               
              'Debug.Print Bvulnerable(StYear, area)
              'Debug.Print Wvul
        Next age   'this loop is from StAge+1
        
'NB: Bmature is missing contribution of Stage so create SB0(Area)
        SB0(StYear, Area) = Bmature(StYear, Area) + R0(Area) * w(StYear, Area, Stage) * FracMat(Stage)
        SBR0(Area) = SB0(StYear, Area) / R0(Area)
        VB0(Area) = Bvulnerable(StYear, Area)
                
        Wvul = 0
        For ilen = iLfull(Area) To Nilens
             Wvul = Wvul + W_L(Area, ilen) * pLage(Area, Stage, ilen)
        Next ilen
        WvulStage(Area) = Wvul
             
    Next Area

    SBR0_avg = 0
    R0total = 0
    For Area = 1 To Nareas
        SBR0_avg = SBR0_avg + R0(Area) * SBR0(Area)
        R0total = R0total + R0(Area)
    Next Area
    
    SBR0_avg = SBR0_avg / R0total
        
   For Area = 1 To Nareas
     SBRXConectividad(Area) = 0
     For i = 1 To Nareas
         'SBRXConectividad(Area) = SBRXConectividad(Area) + Connect(Area, i) * SBR0(i)
         SBRXConectividad(Area) = SBRXConectividad(Area) + Connect(Area, i) * SBR0(i) * R0(i) / R0(Area)
     Next i

       ' Debug.Print SBRXConectividad(Area)
   Next Area
    
  'compute minimum
  minSBRXConectividad = SBRXConectividad(1)

     For Area = 2 To Nareas
          If SBRXConectividad(Area) < minSBRXConectividad Then minSBRXConectividad = SBRXConectividad(Area)
     Next Area
    
  ProdXB = (1 / minSBRXConectividad) * Lambda_ProdXB
    
  Call M6_Prod_Alloc_Larvae.Prod_Alloc_Larvae(StYear, SB0)

  For Area = 1 To Nareas
        
      For year = StYear To StYear + Stage - 1
            Settlers(year, Area) = Settlers(StYear + Stage, Area)
         '      Debug.Print Settlers(year, area)
      Next year
  Next Area

  
  Call Input_Output.Print_Initial_Conditions("Carrying_Capacity")

VB0_all = 0
SB0_all = 0
For Area = 1 To Nareas
    VB0_all = VB0_all + VB0(Area)
    SB0_all = SB0_all + SBR0(Area) * R0(Area)
Next Area

End Sub
'#######################################################################################
'#                            VIRGIN CONDITIONS
'#  Calculate virgin conditions (unharvested equilibrium) by projecting population
'#  forward from carrying capacity. Note that if Lambda < 1 some areas will be limited by
'#  larval supply and will be below K.
'#######################################################################################

Sub Set_Virgin_Conditions()

ReDim Amax(Nareas)
Dim simyear As Integer

' Biomass in StYear is at carrying capacity calculated in Set_Carrying_Capacity

' Simulate with harvest rate 0
        SimEndYear = StYear + 200
    
    'Start from K and simulate forward with no harvest
    For Area = 1 To Nareas
          HRTmp(Area) = 0
    Next Area
                    
    Call Preliminary_Calcs.Initialize_tmp_variables
    
    For simyear = StYear To SimEndYear - 1
               
        Call M4_Calc_Recruits.Deterministic_Recruits(simyear)
            
        Call M6_Prod_Alloc_Larvae.Prod_Alloc_Larvae(simyear, Bmature)
          
        Call M5_Popdyn.PopDyn(simyear)
                        
        Call M2_AnnualUpdate.Annual_update(simyear)
    
     Next simyear
     
'initialize variables for initial condition simulation
     For Area = 1 To Nareas
         For age = Stage To AgePlus
              N(StYear, Area, age) = N(SimEndYear, Area, age)
              mu(StYear, Area, age) = mu(SimEndYear, Area, age)
              sd(StYear, Area, age) = sd(SimEndYear, Area, age)
              w(StYear, Area, age) = w(SimEndYear, Area, age)
          Next age
          
          N(StYear, Area, Stage) = N(SimEndYear - 1, Area, Stage)
          'NB these are missing StAge which is OK
          Btotal(StYear, Area) = Btotal(SimEndYear, Area)
          Bvulnerable(StYear, Area) = Bvulnerable(SimEndYear, Area)
          Bmature(StYear, Area) = Bmature(SimEndYear, Area)
                               
          For year = StYear To StYear + Stage - 1
              Settlers(year, Area) = Settlers(SimEndYear - 1, Area)
          Next year
     Next Area
                        
     Call Input_Output.Print_Initial_Conditions("Virgin_Conditions")

VB0_all = 0
SB0_all = 0
For Area = 1 To Nareas
    VB0_all = VB0_all + VB0(Area)
    SB0_all = SB0_all + SBR0(Area) * R0(Area)
Next Area

End Sub



'#################################################################################
'#                         INITIAL CONDITIONS                                                                         #
'#################################################################################

Sub Set_InitialConditions()
Attribute Set_InitialConditions.VB_ProcData.VB_Invoke_Func = " \n14"
Dim year As Integer, Area As Integer, simyear As Integer

Select Case RunFlags.Initial_Conditions

  Case 1
      
     Call Input_Output.Read_Initial_Conditions("Carrying_Capacity")
     Call Input_Output.Print_Initial_Conditions("Initial_Conditions")
    
  Case 2
      
     Call Input_Output.Read_Initial_Conditions("Virgin_Conditions")
     Call Input_Output.Print_Initial_Conditions("Initial_Conditions")
        
  Case 3
        SimEndYear = StYear + 200
    
    'Start from Virgin and simulate forward under constant harvest rate = HR_start(area)
    '
    For Area = 1 To Nareas
          HRTmp(Area) = HR_start(Area)
    Next Area
                
    For simyear = StYear To SimEndYear - 1
               
        Call M4_Calc_Recruits.Deterministic_Recruits(simyear)
            
        Call M6_Prod_Alloc_Larvae.Prod_Alloc_Larvae(simyear, Bmature)
          
        Call M5_Popdyn.PopDyn(simyear)
                        
        Call M2_AnnualUpdate.Annual_update(simyear)
    
     Next simyear
 
     For Area = 1 To Nareas
         For age = Stage To AgePlus
              N(StYear, Area, age) = N(SimEndYear, Area, age)
              mu(StYear, Area, age) = mu(SimEndYear, Area, age)
              sd(StYear, Area, age) = sd(SimEndYear, Area, age)
              w(StYear, Area, age) = w(SimEndYear, Area, age)
          Next age
          
          N(StYear, Area, Stage) = N(SimEndYear - 1, Area, Stage)  'NB use SimEndYear-1 because after annual update N at Stage is zero
          Btotal(StYear, Area) = Btotal(SimEndYear, Area)
          Bvulnerable(StYear, Area) = Bvulnerable(SimEndYear, Area)
          Bmature(StYear, Area) = Bmature(SimEndYear, Area)
                               
          For year = StYear To StYear + Stage
              Settlers(year, Area) = Settlers(SimEndYear - 1, Area)
          Next year
     Next Area
                        
     Call Input_Output.Print_Initial_Conditions("Initial_Conditions")
  
  Case 4
        'Read from arbitrary input file (e.g. estimated from conditioning)
       
     Call Input_Output.Read_Initial_Conditions("Initial_Conditions")
  
  End Select

End Sub
Sub Rescale_parameters()
Attribute Rescale_parameters.VB_ProcData.VB_Invoke_Func = " \n14"
For Area = 1 To Nareas
        M(Area) = M(Area) / Nt
        k(Area) = k(Area) / Nt
        Rho(Area) = Exp(-k(Area))
Next Area
End Sub

'#################################################################################
'#                INITIALIZE TEMP VARIABLES BEFORE SIMULATION LOOPS                                                                          #
'#################################################################################

Sub Initialize_tmp_variables()

  For Area = 1 To Nareas
        
       For age = Stage To AgePlus
        
           NTmp(Area, age) = N(StYear, Area, age)
           muTmp(Area, age) = mu(StYear, Area, age)
           sdTmp(Area, age) = CVmu(Area) * muTmp(Area, age)
           BvulTmp(Area) = Bvulnerable(StYear, Area)
           BtotTmp(Area) = Btotal(StYear, Area)
          
       Next age
             
  Next Area

End Sub
