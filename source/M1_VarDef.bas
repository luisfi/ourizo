Attribute VB_Name = "M1_VarDef"
'List of Global variables
'This variable allows for Different languages in the model (English and Spanish)

Public Language As String

Public age As Integer, i As Long, j As Long, h As Long, ilen As Integer, Area As Integer

'Declaring dimensioning variables
Public AgePlus As Integer
Public Lfull() As Double
Public AgeFullMature As Integer
Public StYear As Integer
Public EndYear As Integer
Public Nyears As Integer
Public Nareas As Integer
Public Nregions As Integer
Public NBregions As Integer
Public Stage As Integer
Public Nages As Integer
Public GraphFlag As Integer
Public L1 As Double
Public Linc As Double
Public Nilens As Integer
Public Version As String
Public SimEndYear As Integer
Public Ndias_beforeswitch As Integer
'Es el numero de dias minimo que los pescadores se quedan en un lugar una vez que eligen en que area

Public Nt_Season As Integer
Public t_StSeason As Integer
Public t_Repr As Integer
Public FracHRPreRepr As Double
Public Npulses As Integer
Public Sens As Double
Public EffortPulse As Double
Public EffortDistributionFlag As Integer
Public CR() As Double
Public TAC_TAE_HR As Integer
Public MaxEffort As Double
Public Feedback As Boolean
Public TargetHR As Double
Public PulseHR As Double
Public PulseHRadjust As Double
Public TargetSurface As Double
Public RCVirginBiomass_Fraction() As Double
Public RCVirginBiomass_Tolerance() As Double
Public RCPreharvestBiomass_Fraction() As Double
Public RCPreharvestBiomass_Tolerance() As Double
Public RCMinimumDensity() As Double
Public RCMinimumDensity_Tolerance() As Double
Public RCGreaterSize_Fraction() As Double
Public RCGreaterSize_Size() As Double
Public RCGreaterSize_Tolerance() As Double
Public ReOpenConditionFlag As Boolean
Public PartialSurveyFlag As Boolean
Public ReOpenCondition() As Double
Public NOpenConditions As Integer
Public RotationPeriod() As Integer
Public ReOpenConditionValues() As Double
Public RestingTimeFlag As Boolean
Public RestingTime() As Integer
Public AdaptativeRotationFlag As Boolean
Public ShortenTolerance() As Double
Public TAC() As Double
Public TAE() As Double
Public TAC_area() As Double
Public TAE_area() As Double
Public TAC_region() As Double
Public TAE_region() As Double
Public EffortPulseRegion() As Double
Public ClosedArea() As Boolean
Public ClosedAreaTmp() As Boolean
Public ClosedRegion() As Boolean
Public ClosedRegionTmp() As Boolean
Public OpenMonth() As Boolean
Public Bregion() As Integer

'Varies accros Year, Age and Area
Public N() As Double, mu() As Double, sd() As Double

'Varies accros years and areas
Public Bvulnerable() As Double
Public Catch() As Double
Public effort() As Double
Public Larvae() As Double
Public Bmature() As Double
Public Btotal() As Double
Public SurveyAll() As Double
Public SurveyBtot() As Double
Public SurveyBvul() As Double
Public SurveyMat() As Double
Public SurveyNage() As Double
Public SurveyNtot() As Double
Public SurveypL() As Double


'Public SurveyUnit() As Double
'Public SurveyVariable() As Double
Public pLopt As Boolean
Public Nsurveys As Integer
'Public SurveyQ() As Integer
'Public SurveyCV() As Integer
'Public SurveySel() As Integer

Public FracSel() As Double
Public FracSelStAge() As Double
Public FracMat() As Integer
Public Settlers() As Double
Public pL() As Double
Public WvulStage() As Double

'Should I declare them only if conditioning?
Public ObsRec() As Double
Public ObsBvul() As Double
Public ObsAbundance() As Double
Public ObsCatch() As Double
Public NObsBvul As Integer
Public NObsAbundance As Integer

Public TotalLike As Double
Public LikeBvul As Double
Public LikeAbundance As Double
Public LikeRec As Double

'Varies accross areas

Public HR_start() As Double
Public CVmu() As Double
Public q() As Double
Public Region() As Integer
Public Surface() As Double, Lat() As Double, Lon() As Double
Public Atlas() As Double
Public RecMax() As Double
Public EffortTmp() As Double
Public CatchTmp() As Double
Public CumEffort() As Double
Public CatchAdjust() As Double



'Declare carrying capacity and virgin condition (unharvested equilibrium) variables
Public R0() As Double, SBR0() As Double, BR0() As Double, SB0() As Double, VB0() As Double, Alpha0() As Double, Beta0() As Double
Public SBvirgin() As Double, VBvirgin() As Double
Public VB0_all As Double, SB0_all As Double, VBvirgin_all As Double, SBvirgin_all As Double

Public Linf() As Double, aW() As Double, bW() As Double, _
        k() As Double, M() As Double, t0() As Double, _
        Alpha() As Double, _
        Beta() As Double, Rho() As Double, _
        Kcarga() As Double, Kcarga_adults() As Double, Bthreshold() As Double, Btotal_start() As Double, _
        Rmax() As Double, g() As Double, gk() As Double, Bg0() As Double

'Temp variables for intrayear dynamics
Public BvulTmp() As Double, NTmp() As Double, BtotTmp() As Double, muTmp() As Double, sdTmp() As Double, WTmp() As Double, _
HRTmp() As Double
'Fishing dynamics
Public Flag_Rec_Fish() As Integer, Nfracs() As Integer, Z() As Double, frac() As Double, NpulsosMax As Integer, pLage() As Double

'Not categorized yet
Public ProdXB As Double
Public Lambda_ProdXB As Double
Public handling As Double
Public price As Double
Public cost() As Double

Public intfact As Double

Public w() As Double
Public W_L() As Double
Public l() As Double
Public pLageplus() As Double
Public pLStAge() As Double
Public iLfull() As Integer

Public Connect() As Double


Public Type Flags
   Rec As Integer
   Hstrategy As Integer
   RotationType As Integer
   Growth_type As Integer
   Initial_Conditions As Integer
   Run_type As Integer
   VirginAgePlus As Integer

   InputRec As Boolean

   InputBvul As Boolean
   BvulType As Integer
   InputAbundance As Boolean
   AbundanceType As Integer
   InputCatch As Boolean
   Output_Size_W As Boolean
   Output_NAge_NSize As Boolean
   ObsError_Survey As Integer
   ProcError_Rec As Integer
   ProcError_InitConditions As Integer
End Type

Public q_Rec As Double

Public Nt As Integer

Public RunFlags As Flags

Public Candidate_areas() As Integer

Public Nareas_region() As Integer
Public AnnualCatch() As Double


'Monte Carlo
Public Nreplicates As Double
Public RecCV As Double
Public RecTimeCor As Double
Public InitialCV As Double
Public SurveyCV As Double
Public Zvector() As Double
Public N_Zvector As Long
Public iz As Long
Public Rdev() As Double

