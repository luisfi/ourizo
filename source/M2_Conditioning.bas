Attribute VB_Name = "M2_Conditioning"
Option Explicit
Dim PenaltyCatch As Double, ResBvul() As Double, ResAbundance() As Double, year As Integer
Sub Conditioning()
Attribute Conditioning.VB_ProcData.VB_Invoke_Func = " \n14"

    PenaltyCatch = 0

    For year = StYear To EndYear
        
        If RunFlags.InputRec = True Then
            Call M4_Calc_Recruits.Tunned_Recruits(year)
        Else
            Call M4_Calc_Recruits.Deterministic_Recruits(year)
        End If
        
        Call M6_Prod_Alloc_Larvae.Prod_Alloc_Larvae(year)
        
        For Area = 1 To Nareas
            HRTmp(Area) = ObsCatch(year, Area) / BvulTmp(Area)
            
            If HRTmp(Area) <= 0.95 Then
                Catch(year, Area) = ObsCatch(year, Area)
            Else
                HRTmp(Area) = 0.95
                Catch(year, Area) = BvulTmp(Area) * HRTmp(Area)
                PenaltyCatch = PenaltyCatch + 10000
                    'penalty added in case of HRTmp too big, arbitrary 10000
            End If
         Next Area
            
        Call M5_Popdyn.PopDyn(year)
        Call M2_AnnualUpdate.Annual_update(year)
    Next year
End Sub
Sub FitData()
Attribute FitData.VB_ProcData.VB_Invoke_Func = " \n14"
   
   If RunFlags.InputBvul = True Then
   ReDim ResBvul(NObsBvul)
   
        For i = 1 To NObsBvul
            year = ObsBvul(i, 1)
            Area = ObsBvul(i, 2)
            ResBvul(i) = Log(ObsBvul(i, 3)) - Log(Bvulnerable(year, Area) / Surface(Area))
        Next i
           
        If RunFlags.BvulType = 2 Then
            Dim log_q_Bvul As Double
            log_q_Bvul = 0
            For i = 1 To NObsBvul
               log_q_Bvul = log_q_Bvul + Log(ObsBvul(i, 3)) - Log(Bvulnerable(year, Area) / Surface(Area))
            Next i
            log_q_Bvul = log_q_Bvul / NObsBvul
            
            For i = 1 To NObsBvul
                ResBvul(i) = ResBvul(i) - log_q_Bvul
            Next i
        
        End If
   
   End If
   
   If RunFlags.InputAbundance = True Then
      ReDim ResAbundance(NObsAbundance)
      Dim Nvulnerable As Double
      For i = 1 To NObsAbundance
        year = ObsAbundance(i, 1)
        Area = ObsAbundance(i, 2)
        
        Nvulnerable = 0
        For age = 1 To AgePlus
            Nvulnerable = Nvulnerable + n(year, Area, age) * FracSel(Area, age)
        Next age
        ResAbundance(i) = Log(ObsAbundance(i, 3)) - Log(Nvulnerable / Surface(Area))
      Next i
      
        If RunFlags.AbundanceType = 2 Then
            Dim log_q_Abundance As Double
            log_q_Abundance = 0
            For i = 1 To NObsAbundance
               log_q_Abundance = log_q_Abundance + ResAbundance(i)
            Next i
            log_q_Abundance = log_q_Abundance / NObsAbundance
            
            For i = 1 To NObsAbundance
                ResAbundance(i) = ResAbundance(i) - log_q_Abundance
            Next i

        End If
   End If


'START NEW PRINT FOR TUNING
'Erase tuning from previous run

Sheets("OutTuning").Activate
Columns("A:K").Select
Selection.ClearContents

'Print labels
Worksheets("OutTuning").Rows(1).Columns(1) = "Year"
Worksheets("OutTuning").Rows(1).Columns(2) = "Area"
Worksheets("OutTuning").Rows(1).Columns(3) = "Region"
Worksheets("OutTuning").Rows(1).Columns(4) = "Recruits"
Worksheets("OutTuning").Rows(1).Columns(5) = "ObsRec"
Worksheets("OutTuning").Rows(1).Columns(6) = "Bvulnerable"
Worksheets("OutTuning").Rows(1).Columns(7) = "ObsBvul"
Worksheets("OutTuning").Rows(1).Columns(8) = "Abundance"
Worksheets("OutTuning").Rows(1).Columns(9) = "ObsAbundance"


'Print year, area, region
For year = 1 To Nyears
    For Area = 1 To Nareas
        Worksheets("OutTuning").Rows(year + 1 + (Nyears) * (Area - 1)).Columns(1) = StYear - 1 + year
        Worksheets("OutTuning").Rows(year + 1 + (Nyears) * (Area - 1)).Columns(2) = Area
        Worksheets("OutTuning").Rows(year + 1 + (Nyears) * (Area - 1)).Columns(3) = Region(Area)
    Next Area
Next year

'If tuning to recruits
If RunFlags.InputRec = True Then
    For year = 1 To Nyears
        For Area = 1 To Nareas
            Worksheets("OutTuning").Rows(year + 1 + (Nyears) * (Area - 1)).Columns(4) = (ObsRec(StYear - 1 + year, Area) * q_Rec / Exp(Rdev(StYear - 1 + year, Area))) * Exp(0.5 * RecCV ^ 2)
            Worksheets("OutTuning").Rows(year + 1 + (Nyears) * (Area - 1)).Columns(5) = ObsRec(StYear - 1 + year, Area)
        Next Area
    Next year
End If


'If tuning to Bvulnerable
If RunFlags.InputBvul = True Then
    For year = 1 To Nyears
        For Area = 1 To Nareas
            Worksheets("OutTuning").Rows(year + 1 + (Nyears) * (Area - 1)).Columns(6) = Bvulnerable(StYear - 1 + year, Area) / Surface(Area)
        Next Area
    Next year
      
    For i = 1 To NObsBvul
        year = ObsBvul(i, 1)
        Area = ObsBvul(i, 2)
        Worksheets("OutTuning").Rows(year - StYear + 2 + (Nyears) * (Area - 1)).Columns(7) = ObsBvul(i, 3) / Exp(log_q_Bvul)
    Next i

End If

'If tuning to Abundance
If RunFlags.InputAbundance = True Then
    For year = 1 To Nyears
        For Area = 1 To Nareas
            Nvulnerable = 0
            For age = 1 To AgePlus
                Nvulnerable = Nvulnerable + n(StYear - 1 + year, Area, age) * FracSel(Area, age)
            Next age
            Worksheets("OutTuning").Rows(year + 1 + (Nyears) * (Area - 1)).Columns(8) = Nvulnerable / Surface(Area)
        Next Area
    Next year

    For i = 1 To NObsAbundance
        year = ObsAbundance(i, 1)
        Area = ObsAbundance(i, 2)
        Worksheets("OutTuning").Rows(year - StYear + 2 + (Nyears) * (Area - 1)).Columns(9) = ObsAbundance(i, 3) / Exp(log_q_Abundance)
    Next i

End If

End Sub

Sub CalcLikelihood()
Attribute CalcLikelihood.VB_ProcData.VB_Invoke_Func = " \n14"
   Dim zeta As Double, epsCV As Double
    
    TotalLike = 0
   
   If RunFlags.InputBvul = True Then
      LikeBvul = 0
      
      For i = 1 To NObsBvul
         LikeBvul = LikeBvul + ResBvul(i) ^ 2 / ObsBvul(i, 4)
      Next i
      TotalLike = TotalLike + 0.5 * LikeBvul
   End If
   
   If RunFlags.InputAbundance = True Then
      LikeAbundance = 0
      
      For i = 1 To NObsAbundance
         LikeAbundance = LikeAbundance + ResAbundance(i) ^ 2 / ObsAbundance(i, 4)
      Next i
      TotalLike = TotalLike + 0.5 * LikeAbundance
      
   End If
   
   If RunFlags.InputRec = True Then
      LikeRec = 0
      epsCV = RecCV * (1 - RecTimeCor ^ 2)
      
      'esto solo asume autocorrelation temporal- falta incorporar espacial
      
      For Area = 1 To Nareas
       LikeRec = LikeRec + Rdev(StYear, Area) ^ 2 / RecCV
       For year = StYear + 1 To EndYear
         zeta = (Rdev(year, Area) - RecTimeCor * Rdev(year - 1, Area)) / epsCV
         LikeRec = LikeRec + zeta ^ 2
       Next year
      Next Area
      TotalLike = TotalLike + 0.5 * LikeRec
      
   End If
   
   TotalLike = TotalLike + PenaltyCatch
   
   
'Print Likelihoods
   
Worksheets("OutTuning").Rows(1).Columns(10) = "LikeType"
Worksheets("OutTuning").Rows(1).Columns(11) = "LikeValue"
Worksheets("OutTuning").Rows(2).Columns(10) = "LikeRec"
Worksheets("OutTuning").Rows(2).Columns(11) = LikeRec
Worksheets("OutTuning").Rows(3).Columns(10) = "LikeBvul"
Worksheets("OutTuning").Rows(3).Columns(11) = LikeBvul
Worksheets("OutTuning").Rows(4).Columns(10) = "LikeAbundance"
Worksheets("OutTuning").Rows(4).Columns(11) = LikeAbundance
Worksheets("OutTuning").Rows(5).Columns(10) = "PenaltyCatch"
Worksheets("OutTuning").Rows(5).Columns(11) = PenaltyCatch
Worksheets("OutTuning").Rows(6).Columns(10) = "TotalLike"
Worksheets("OutTuning").Rows(6).Columns(11) = TotalLike
   
End Sub
