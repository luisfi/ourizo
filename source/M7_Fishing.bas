Attribute VB_Name = "M7_Fishing"
Option Explicit
Option Base 1

Sub Fishing(year, t)
Attribute Fishing.VB_ProcData.VB_Invoke_Func = " \n14"
'This subrutine's main output is to calculate HRTmp y Catch

Dim area As Integer, pulse As Integer, i_t As Integer, i_area As Integer, rr As Integer, _
    j As Integer, Nopenareas As Integer, Nopenregions As Integer, Nfishedareas As Integer, _
    NN As Integer, i_rr As Integer
Dim Max As Double
Dim profit() As Double, CatchPulseRegion() As Double

ReDim EffortTmp(Nareas) As Double, CatchTmp(Nareas) As Double, CR(Nareas) As Double
ReDim CatchAdjust(Nregions) As Double, CatchPulseRegion(Nregions) As Double
ReDim profit(Nareas)

Dim IDopenareaTmp() As Integer, IDfishedarea() As Integer


Select Case RunFlags.Hstrategy

Case 1 ' ROTATIONAL ESTO ES ANUAL
 
  For area = 1 To Nareas
     If ClosedAreaTmp(area) = False Then
        If TAC_TAE_HR = 1 Then
           Catch(year, area) = TAC_area(year, area)
        ElseIf TAC_TAE_HR = 2 Then
           MsgBox ("Rotation scheme not implemented for additional effort (input) control")
           End
        ElseIf TAC_TAE_HR = 3 Then
           Catch(year, area) = BvulTmp(area) * TargetHR
        Else
           HRTmp(area) = PulseHR * PulseHRadjust
        
           'Update Atlas for the harvested areas
           If PartialSurveyFlag = True Then Atlas(area) = SurveyBvul(1, year, area) * (1 - HRTmp(area))

           Catch(year, area) = BvulTmp(area) * HRTmp(area)
        End If
     Else
        HRTmp(area) = 0
     End If
  Next area
   
Case 2  'SPATIAL MANAGEMENT BY INDIVIDUAL AREA

' set ID of areas open to fishing
 
  For area = 1 To Nareas
    If ClosedAreaTmp(area) = False Then
             
        If TAC_TAE_HR = 1 Then
            
            HRTmp(area) = TAC_area(year, area) / BvulTmp(area)
            
                If HRTmp(area) <= 0.9 Then
                        Catch(year, area) = TAC_area(year, area)
                Else
                        HRTmp(area) = 0.9
                        Catch(year, area) = BvulTmp(area) * HRTmp(area)
                End If
                        
        ElseIf TAC_TAE_HR = 2 Then
                    
            HRTmp(area) = 1 - Exp(-q(area) * TAE_area(year, area))
            Catch(year, area) = BvulTmp(area) * HRTmp(area)
                
        ElseIf TAC_TAE_HR = 3 Then
        
            HRTmp(area) = TargetHR
            Catch(year, area) = BvulTmp(area) * HRTmp(area)
            
        End If
                
    Else      'the area is closed
        HRTmp(area) = 0
    End If
  Next area

Case 3 ' GLOBAL or REGIONAL MANAGEMENT (implica que hay que distribuir el esfuerzo entre areas e.g. IFD)
    
    Dim IDopenregionTmp() As Integer
        
        i_rr = 0
        For rr = 1 To Nregions
            If ClosedRegionTmp(rr) = False Then
                i_rr = i_rr + 1
                ReDim Preserve IDopenregionTmp(i_rr)
                IDopenregionTmp(i_rr) = rr
            End If
        Next rr
        Nopenregions = i_rr
        
        NN = Nopenregions   'NN is number of open regions at the start of the month
                    
        For area = 1 To Nareas
           EffortTmp(area) = 0
        Next area
                      
      
    If EffortDistributionFlag = 1 Then     ' Subcaso 1 IFD-  This is the obsolete ideal free distribution algorithm
                                    
        Dim EffortPulseArea() As Double
        ReDim EffortPulseArea(Nareas)
        Dim CatchPulseArea() As Double
        ReDim CatchPulseArea(Nareas)
        
        For pulse = 1 To Npulses 'Npulses is number of fishing pulses (when effort can be distributed) whitin the pop dyn. eg is 4 when they stay at least a week on a ground and the pop dyn is monthly.
                    ' set ID of areas open to fishing
                        
              For rr = 1 To Nregions
                  CatchPulseRegion(rr) = 0
              Next rr
                  
           Select Case TAC_TAE_HR
           
           Case 1             'TAC by region
                              'NB!:  assumes that effort is global, not regionalized
                  i_area = 0
                  Max = 0
                  For area = 1 To Nareas
                     If ClosedAreaTmp(area) = False Then
                         i_area = i_area + 1
                         ReDim Preserve IDopenareaTmp(i_area)
                         IDopenareaTmp(i_area) = area
                         CR(area) = BvulTmp(area) * q(area)
                           'Determination of area with highest catch rate
                         Max = 1 / 2 * (CR(area) + Max + Abs(Max - CR(area)))
                     End If
                  Next area
                        
                  Nopenareas = i_area
                        
                  'Determine number of areas to fish
                  Nfishedareas = 0
                  For i_area = 1 To Nopenareas
                      area = IDopenareaTmp(i_area)
                      If ((1 - CR(area) / Max) < Sens) = True Then
                        Nfishedareas = Nfishedareas + 1
                        ReDim Preserve IDfishedarea(Nfishedareas)
                        IDfishedarea(Nfishedareas) = area
                      End If
                  Next i_area
                        
                  For i_area = 1 To Nfishedareas   'Allocation of effort
                      area = IDfishedarea(i_area)
                      EffortPulseArea(area) = EffortPulse / Nfishedareas    'EffortPulse set by MP
                      CatchPulseArea(area) = BvulTmp(area) * (1 - Exp(-EffortPulseArea(area) * q(area)))
                      rr = Region(area)
                      CatchPulseRegion(rr) = CatchPulseRegion(rr) + CatchPulseArea(area)
                  Next i_area
                                     
                  For i_rr = 1 To Nopenregions   'loop over regions open AT BEGINING of month even though some
                                                       'may be already closed. Doesn't matter because their CatchPulseRegion(rr) =0
                        rr = IDopenregionTmp(i_rr)
                        If ClosedRegionTmp(rr) = False Then
                           If (AnnualCatch(rr) + CatchPulseRegion(rr) > TAC_region(rr, year)) = True Then
                               CatchAdjust(rr) = (TAC_region(rr, year) - AnnualCatch(rr)) / CatchPulseRegion(rr) 'This is going to be < 1
                               ClosedRegionTmp(rr) = True
                               NN = NN - 1           'decrease number of open regions for each one closed during this pulse
                               For i_area = 1 To Nareas_region(rr)
                                  area = Candidate_areas(rr, i_area)
                                  CatchPulseArea(area) = CatchPulseArea(area) * CatchAdjust(rr)
                                  EffortPulseArea(area) = -Log(1 - CatchPulseArea(area) / BvulTmp(area)) / q(area)                           'Rough approximate
                                  ClosedAreaTmp(area) = True
                               Next i_area
                           End If
                           AnnualCatch(rr) = AnnualCatch(rr) + CatchPulseRegion(rr)
                        End If
                  Next i_rr
                        
                  For i_area = 1 To Nfishedareas
                         area = IDfishedarea(i_area)
                         rr = Region(area)
                         EffortTmp(area) = EffortTmp(area) + EffortPulseArea(area)   'cumulative over month
                         BvulTmp(area) = BvulTmp(area) - CatchPulseArea(area)
                         Catch(year, area) = Catch(year, area) + CatchPulseArea(area)
                  Next i_area
                                                                 
                  If NN = 0 Then     'no more regions open to fishing
                        For i_t = t + 1 To Nt
                               OpenMonth(year, i_t) = False
                        Next i_t
                        Exit For     'sale del loop de Npulses
                  End If
                      
           Case 2     'TAE by region
               
             For i_rr = 1 To Nopenregions
                       
                rr = IDopenregionTmp(i_rr)
                    
                Max = 0
                For i_area = 1 To Nareas_region(rr)
                    area = Candidate_areas(rr, i_area)
                    CR(area) = BvulTmp(area) * q(area)
                      'Determination of area with highest catch rate
                    Max = 1 / 2 * (CR(area) + Max + Abs(Max - CR(area)))
                Next i_area
               
                'Determination of number of areas to fish
                Nfishedareas = 0
                
                For i_area = 1 To Nareas_region(rr)
                    area = Candidate_areas(rr, i_area)
                    If ((1 - CR(area) / Max) < Sens) = True Then
                         Nfishedareas = Nfishedareas + 1
                         ReDim Preserve IDfishedarea(Nfishedareas)
                         IDfishedarea(Nfishedareas) = area
                    End If
                Next i_area
                   
                For i_area = 1 To Nfishedareas   'Allocate effort
                      area = IDfishedarea(i_area)
                      EffortPulseArea(area) = EffortPulseRegion(rr) / Nfishedareas
                      CatchPulseArea(area) = BvulTmp(area) * (1 - Exp(-EffortPulseArea(area) * q(area)))
                      EffortTmp(area) = EffortTmp(area) + EffortPulseArea(area)   'cumulative over month
                      BvulTmp(area) = BvulTmp(area) - CatchPulseArea(area)
                      Catch(year, area) = Catch(year, area) + CatchPulseArea(area)
                Next i_area
                
              Next i_rr
                     
           End Select
           
        Next pulse
     
        For area = 1 To Nareas
           effort(year, area) = effort(year, area) + EffortTmp(area)
           HRTmp(area) = (1 - Exp(-EffortTmp(area) * q(area)))
        Next area
     
    ElseIf EffortDistributionFlag = 2 Then      'IFD modified from Walters & Martel
               
        Dim ordenvector() As Double
        Dim orderedareas() As Integer
        Dim pr0 As Double
        Dim proci As Double
        Dim Ninitial As Double
                     
                
      Select Case TAC_TAE_HR
        
        Case 1  'TAC by region
         
         Dim TACleft() As Double
         ReDim TACleft(Nopenregions)
         Dim iregion() As Integer
         ReDim iregion(Nregions)
         Dim boundbyTAC As Boolean
         
         For i_rr = 1 To Nopenregions
            rr = IDopenregionTmp(i_rr)
            iregion(rr) = i_rr
            TACleft(i_rr) = TAC_region(rr, year) - AnnualCatch(rr)
         Next i_rr
         
         i_area = 0
         For area = 1 To Nareas
            If ClosedAreaTmp(area) = False Then
               i_area = i_area + 1
               ReDim Preserve ordenvector(i_area)
               ReDim Preserve orderedareas(i_area)
               CR(area) = q(area) * BvulTmp(area) / (1 + q(area) * handling * BvulTmp(area))
               profit(area) = price * CR(area) - cost(area)
               ordenvector(i_area) = profit(area)
               orderedareas(i_area) = area
            End If
         Next area
         Nopenareas = i_area
            
         Call order(ordenvector, orderedareas)   'sort areas by profitability
        
       
         If Nopenregions = 1 Then
                
           'Calculate final values for pr0 (profitability at which fished areas are equalized)
           'so that TACleft and EffortPulse (max effort) are not exceeded and number of fished areas
           ' calls Newton-Rabson from subroutine
            
            Call EqualizePops1TAC(orderedareas, profit, Nopenareas, EffortPulse, TACleft(1), pr0, Nfishedareas, boundbyTAC)
                       
            For i = 1 To Nfishedareas
               area = orderedareas(i)
               Call calcAll(pr0, area, year)
            Next i
            
            If boundbyTAC = True Then     'reached TAC
              For i_t = t + 1 To Nt
                 OpenMonth(year, i_t) = False
              Next i_t
            End If
    
         
         ElseIf Nopenregions > 1 Then
         
           
           Call EqualizePopsTACregion(orderedareas, profit, Nopenareas, Nopenregions, IDopenregionTmp, iregion, EffortPulse, TACleft, year, boundbyTAC)
            
           
           'aqui hay que chequear si todas las regiones estan listas
           If boundbyTAC = True Then     'TAC reached in all regions
              For i_t = t + 1 To Nt
                 OpenMonth(year, i_t) = False
              Next i_t
            End If
           
         End If   'finish Nregions > 1
         
   
        Case 2  'TAE by region
        
         For i_rr = 1 To Nopenregions
           
            rr = IDopenregionTmp(i_rr)
                       
            ReDim ordenvector(1 To Nareas_region(rr))
            ReDim orderedareas(1 To Nareas_region(rr))
                    
            For i_area = 1 To Nareas_region(rr)
                 area = Candidate_areas(rr, i_area)
                 CR(area) = q(area) * BvulTmp(area) / (1 + q(area) * handling * BvulTmp(area))
                 profit(area) = price * CR(area) - cost(area)
                 ordenvector(i_area) = profit(area)
                 orderedareas(i_area) = area
            Next i_area
            
            Call order(ordenvector, orderedareas)

            Nfishedareas = EqualizePops(orderedareas, profit, BvulTmp, Nareas_region(rr), EffortPulseRegion(rr))

           'Calcula initial values for pr0 (profitability at which fished areas are equalized)
            If Nfishedareas = Nareas_region(rr) Then
                pr0 = 0.8 * profit(orderedareas(Nfishedareas))  'initialize pro a bit lower than the last area fished
            Else
                pr0 = (profit(orderedareas(Nfishedareas)) + profit(orderedareas(Nfishedareas + 1))) / 2 'initialize pro at an intermediate value between last fished area and first unfished
            End If
        
            pr0 = calcPr0ET(pr0, orderedareas, Nfishedareas, EffortPulseRegion(rr))
 
            'Calculates efforts and fished abundances as a function of pro
            For i = 1 To Nfishedareas
               area = orderedareas(i)
               Call calcAll(pr0, area, year)
            Next
         Next i_rr
         
        End Select
    ElseIf EffortDistributionFlag = 3 Then   'GRAVITACIONAL
        MsgBox ("Gravitational model of effort allocation not implemented - Need to change flag in Fishing Effort Distribution")
        End
       
              '   For i_area = 1 To Nopenareas   'Allocation of effort
             '       area = IDopenareas(i_area)
             '
             '           EffortTmp(area) = EffortPulse * CR(area) / CR_all
             '
             '           CatchTmp(area) = BvulTmp(area) * (1 - Exp(-EffortTmp(area) * Q(area)))
             '           rr = Region(area)
             '
             '           catchpulse(rr) = catchpulse(rr) + CatchTmp(area)
                                        
             '           If TAC_TAE_HR = 1 Then
             '
             '               If catchpulse(rr) > TAC_region(rr, year) Then
             '               'Adjust catch and remaining effort
             '
             '           End If
                        
                        
'                Next i_area
'
    End If
 End Select
 End Sub
 
Public Function EqualizePops(orden, profit, Ninitial, npops, ET)
'calculates effort needed to equalize
  Dim eff As Double
  Dim Nfin As Double
  Dim proci As Double
  Dim indb As Integer
  Dim ind As Integer
  
    j = 1

    Do While eff < ET And j < npops

       j = j + 1
       eff = 0   'total effort in this pulse is computed at the end of loop
       indb = orden(j)   'this is the actual Area
       For i = 1 To j - 1 'Calcula los elementos del EP
          ind = orden(i)
          proci = profit(indb) + cost(ind)
          Nfin = proci / (price * q(ind) - q(ind) * handling * proci)
          eff = eff + (handling * (Ninitial(ind) - Nfin) - Log(Nfin / Ninitial(ind)) / q(ind))
       Next
 
    Loop
'returns the number of equalized areas
    If eff > ET Then
      EqualizePops = j - 1
    Else
      EqualizePops = j
    End If
      
End Function
Sub EqualizePops1TAC(orderedareas, profit, npops, ET, TACtmp, pr0, Nfishedareas, boundbyTAC)
'calculates effort needed to equalize
  Dim eff As Double
  Dim cat As Double
  Dim Nfin As Double
  Dim Ninitial As Double
  Dim proci As Double
  Dim indb As Integer
  Dim ind As Integer
  Dim j As Integer

  Dim boundbyET As Boolean
  Dim bound As Boolean
  Dim i_t As Integer
  
  boundbyTAC = False
  boundbyET = False
  bound = False
  
  
    For j = 2 To npops
       eff = 0
       cat = 0
       indb = orderedareas(j)   'this is the actual Area
       For i = 1 To j - 1 'Calcula los elementos del EP
          ind = orderedareas(i)
          proci = profit(indb) + cost(ind)
          Ninitial = BvulTmp(ind)
          Nfin = proci / (price * q(ind) - q(ind) * handling * proci)
          cat = cat + Ninitial - Nfin
          eff = eff + (handling * (Ninitial - Nfin) - Log(Nfin / Ninitial) / q(ind))
       Next i
 
       If (eff > ET) Then
          boundbyET = True
          bound = True
          If (cat > TACtmp) Then boundbyTAC = True
          Exit For
       ElseIf (cat > TACtmp) Then
          bound = True
          boundbyTAC = True
          Exit For
       End If
    Next j
   
   'number of equalized areas
    If bound = True Then
       Nfishedareas = j - 1
       pr0 = (profit(orderedareas(Nfishedareas)) + profit(orderedareas(Nfishedareas + 1))) / 2 'initialize pro at an intermediate value between last fished area and first unfished
      
       If boundbyET = True Then
          If boundbyTAC = False Then
             pr0 = calcPr0ET(pr0, orderedareas, Nfishedareas, ET)
          Else
             pr0 = calcPr0TACandET(pr0, orderedareas, Nfishedareas, TACtmp, ET, boundbyTAC)
          End If
    
       ElseIf boundbyTAC = True Then   'not bounded by effort
          pr0 = calcPr0TAC(pr0, orderedareas, Nfishedareas, TACtmp)
       End If
    Else   'fish all areas
      
        Nfishedareas = npops
        pr0 = 0.9 * profit(orderedareas(Nfishedareas))  'initialize pro a bit lower than the last area fished
        pr0 = calcPr0TACandET(pr0, orderedareas, Nfishedareas, TACtmp, ET, boundbyTAC)
   
   End If
   
End Sub
Sub EqualizePopsTACregion(orderedareas, profit, npops, Nopenregions, IDopenregionTmp, iregion, ET, TACleft, year, boundbyTAC)
      
'calculates effort needed to equalize profit and meet TAC by region
  Dim eff As Double
  Dim effarea As Double
  Dim cat As Double
  Dim Nfin As Double
  Dim Ninitial As Double
  Dim proci As Double
  Dim indb As Integer
  Dim j As Integer
  Dim i_rr As Integer
  Dim i_area As Integer
  Dim rr As Integer
  Dim ii As Integer
  Dim Nopen As Integer
  Dim ETleft As Double
  Dim pr0 As Double
  Dim NN As Integer
  Dim Nfishedareas As Integer
  Dim fishedareas() As Integer
  Dim pr0region() As Double
  ReDim pr0region(Nopenregions)
  Dim Nfishedareasregion() As Integer
  ReDim Nfishedareasregion(Nopenregions)
  Dim i_t As Integer
  
  ReDim EffortPulseRegion(Nopenregions)
  
  ETleft = ET     'remaining available effort after some regions reach their TAC
  Nopen = Nopenregions
  
  j = 1
  eff = 0
  
  Do While eff < ETleft And j < npops And ETleft > 0
      j = j + 1
      cat = 0
      For i_rr = 1 To Nopenregions
         rr = IDopenregionTmp(i_rr)
         If (ClosedRegionTmp(rr) = False) Then 'NB: a region may be closed if it
                              'reached its TAC at a higher pro (smaller j in do while loop) level
            Nfishedareasregion(i_rr) = 0  'this is numer of fished areas in this pulse e.g. month)
            CatchPulseRegion(i_rr) = 0
            EffortPulseRegion(i_rr) = 0
         End If
      Next i_rr
      
      eff = 0
      indb = orderedareas(j)   'this is the actual Area
      For i = 1 To j - 1
          area = orderedareas(i)   'area fished
       
          If (ClosedAreaTmp(area) = False) Then
             rr = Region(area)
             i_rr = iregion(rr)
             Nfishedareasregion(i_rr) = Nfishedareasregion(i_rr) + 1
             proci = profit(indb) + cost(area)
             Nfin = proci / (price * q(area) - q(area) * handling * proci)
             cat = BvulTmp(area) - Nfin
             
             If (CatchPulseRegion(i_rr) + cat > TACleft(i_rr)) Then
                
                pr0 = profit(indb)  'use this as initial value for Newton-R
                                
                NN = Nfishedareasregion(i_rr)
                ReDim fishedareas(NN)
                ii = 0
                For i_area = 1 To j - 1 'select from all fished areas those that are in the i_rr-th open region
                   area = orderedareas(i_area)
                   If (Region(area) = rr) Then
                      ii = ii + 1
                      fishedareas(ii) = area
                   End If
                Next i_area
                
                pr0 = calcPr0TACandET(pr0, fishedareas, NN, TACleft(i_rr), ETleft, boundbyTAC)
                If (boundbyTAC = True) Then
                   ClosedRegionTmp(rr) = True  'cerrar la region y las areas de esta region
                   For i_area = 1 To Nareas_region(rr)
                      area = Candidate_areas(rr, i_area)
                      ClosedAreaTmp(area) = True
                   Next i_area
                   Nopen = Nopen - 1
                End If
                
                For ii = 1 To NN
                   area = fishedareas(ii)
                   Call calcAll(pr0, area, year)
                   ETleft = ETleft - EffortTmp(area)
                Next ii
             Else
                Ninitial = BvulTmp(area)
                effarea = handling * cat - Log(Nfin / Ninitial) / q(area)
                CatchPulseRegion(i_rr) = CatchPulseRegion(i_rr) + cat
                eff = eff + effarea
             End If
          End If
      Next i
     Loop
    
    If Nopen = 0 Then
            
            boundbyTAC = True
            
    ElseIf Nopen = 1 Then
        
       Dim subset() As Integer
       Dim iopenregion As Integer
       
        For i_rr = 1 To Nopenregions
        rr = IDopenregionTmp(i_rr)
          If (ClosedRegionTmp(rr) = False) Then
            iopenregion = i_rr
            Exit For
          End If
        Next i_rr
        
        ReDim subset(Nareas_region(rr))

        
        ii = 0
        For i_area = 1 To npops
          If (Region(i_area) = rr) Then
              ii = ii + 1
              subset(ii) = orderedareas(i_area)
          End If
        Next i_area
        
        Call EqualizePops1TAC(subset, profit, Nareas_region(rr), ETleft, TACleft(iopenregion), pr0, Nfishedareas, boundbyTAC)
                       
            For i = 1 To Nfishedareas
               area = subset(i)
               Call calcAll(pr0, area, year)
            Next i
            
    ElseIf eff > ETleft And ETleft > 0 And Nopen > 1 Then
     
     'if effort in regions that did not reach their TAC exceeded ETleft, need to equalize profit at a higher value. Number of fished areas
     'is still correct at j-1 because ET was not exceeded when all areas were leveled at the (j-1)-th profit.
     'Fishing will be bound by ET so all areas in regions still open can be
     'collected before calling the N-R algorithm
        boundbyTAC = False
        Nfishedareas = 0   'calculate number of areas to fish in regions that have not
                            'reached their TAC
        For i_rr = 1 To Nopenregions
           rr = IDopenregionTmp(i_rr)
           If ClosedRegionTmp(rr) = False Then    'pr0 for regions that are closed already do not need to be modified
              Nfishedareas = Nfishedareas + Nfishedareasregion(i_rr)
           End If
        Next i_rr
        
        ReDim fishedareas(Nfishedareas)
        ii = 0
        For i_area = 1 To j - 1 'collect from all fished areas those that have not reached their region TAC
            area = orderedareas(i_area)
            If (ClosedAreaTmp(area) = False) Then
                 ii = ii + 1
                 fishedareas(ii) = area
            End If
        Next i_area
        
        pr0 = profit(orderedareas(j - 1))   'initial value for NR
                
        pr0 = calcPr0ET(pr0region(i_rr), fishedareas, Nfishedareas, ETleft)
        'these areas remain open because they are bound by effort, not TAC
        For i_area = 1 To Nfishedareas
           area = fishedareas(i_area)
           Call calcAll(pr0, area, year)
        Next i_area
    Else  'nopen > 1 and ETleft not exceeded after completing loop means that j = npops
          'and the nopen regions that have not reached their TAC can be fished at a lower pro
          'Start reducing pro from lowest

        Dim pr0min As Double
        Dim pr0max As Double
        
        pr0max = profit(orderedareas(npops))
        
        Nfishedareas = 0   'calculate number of areas to fish in regions that have not
                            'reached their TAC
       
        ReDim fishedareas(1)
        Nfishedareas = 0
   
        For i_area = 1 To npops  'collect from all fished areas those that have not reached their region TAC
            area = orderedareas(i_area)
            If (ClosedAreaTmp(area) = False) Then
                 Nfishedareas = Nfishedareas + 1
                 ReDim Preserve fishedareas(Nfishedareas)
                 fishedareas(Nfishedareas) = area
            End If
        Next i_area
        
        pr0min = calcPr0ET(pr0max, fishedareas, Nfishedareas, ETleft)
        
        ReDim IDopenregion(Nopen)
        Dim i_rr2 As Integer
        
        i_rr2 = 0   'collect regions opened
    
        For i_rr = 1 To Nopenregions
        rr = IDopenregion(i_rr)
          If (ClosedRegionTmp(rr) = False) Then
             i_rr2 = i_rr2 + 1
             IDopenregion(i_rr2) = iregion(rr)
             TACleft(i_rr2) = TACleft(i_rr)
          End If
        Next i_rr
        ReDim Preserve TACleft(Nopen)
           
     
     'calculate pr0 at which TAC would be reached in each region assuming all areas fished
        For i_rr = 1 To Nopen
           rr = IDopenregion(i_rr)
           For i_area = 1 To Nareas_region(rr)
              fishedareas(i_area) = Candidate_areas(rr, i_area)
           Next i_area
          
           pr0region(i_rr) = calcPr0TAC(pr0max, fishedareas, Nareas_region(rr), TACleft(i_rr))
        
        Next i_rr
        
        Call order(pr0region, IDopenregion)
        
        If pr0min >= pr0region(1) Then
           'fish all areas up to promin  and do not close any area because they are bound by ETleft
           For i_rr = 1 To Nopen
              rr = IDopenregion(i_rr)
              For i_area = 1 To Nareas_region(rr)
                 
                 area = Candidate_areas(rr, i_area)
                
                 Call calcAll(pr0min, area, year)
              
              Next i_area
           Next i_rr
           boundbyTAC = False
           
        Else   'enough effort left for at least some areas to reach TAC
          
          pr0 = pr0region(1)
          
          For i_rr = 1 To Nopen
          
              pr0 = pr0region(i_rr)
              
              If (pr0 >= pr0min) Then
                 Exit For
              End If
              rr = IDopenregion(i_rr)
              ClosedRegionTmp(rr) = True
                
              For i_area = 1 To Nareas_region(rr)
                 area = Candidate_areas(rr, i_area)
                 ClosedAreaTmp(area) = True
                 Call calcAll(pr0, area, year)
                 ETleft = ETleft - EffortTmp(area)
              Next i_area
          Next i_rr
          
          Dim region1 As Integer
          
          region1 = i_rr   'first region in which pr0 for reaching TAC < pr0min (boundbyETleft)
        
          For i_rr = region1 To Nopen
             pr0 = pr0region(i_rr)
       
       'calc total effort to level remaining areas to pr0
             eff = 0
             For i_rr2 = i_rr To Nopen
                rr = IDopenregion(i_rr2)
                For i_area = 1 To Nareas_region(i_rr2)
                    area = Candidate_areas(rr, i_area)
                    eff = eff + getE(pr0, area)
                Next i_area
             Next i_rr2
           
             If (eff <= ETleft) Then   'close area i_rr boundbyTAC and continue lowering pr0
                rr = IDopenregion(i_rr)
                ClosedRegionTmp(rr) = True
                
                For i_area = 1 To Nareas_region(rr)
                   area = Candidate_areas(rr, i_area)
                   ClosedAreaTmp(area) = True
                   Call calcAll(pr0, area, year)
                   ETleft = ETleft - EffortTmp(area)
                Next i_area
              
             Else 'all remaining regions bound by ETleft
               
                boundbyTAC = False
                Nfishedareas = 0
                ReDim fishedareas(1)
                For i_rr2 = region1 To Nopen
                   rr = IDopenregion(i_rr2)
                   For i_area = 1 To Nareas_region(rr)
                       area = Candidate_areas(rr, i_area)
                       Nfishedareas = Nfishedareas + 1
                       ReDim Preserve fishedareas(Nfishedareas)
                       fishedareas(Nfishedareas) = area
                   Next i_area
                Next i_rr2
               
                pr0 = calcPr0ET(pr0, fishedareas, Nfishedareas, ETleft)
              
                For i_rr2 = region1 To Nopen
                   rr = IDopenregion(i_rr2)
                   For i_area = 1 To Nareas_region(rr)
                       area = Candidate_areas(rr, i_area)
                       Call calcAll(pr0, area, year)
                   Next i_area
                Next i_rr2
                Exit For
             End If
            
          Next i_rr   'continue lowering pr0
        End If
    End If
End Sub

Function calcPr0ET(pr0, sortedfishedareas, Nfishedareas, ET)
'Función que estima la productividad a la que convergen las poblaciones en la Ideal Free Distribution
'Para casos en los que el esfuerzo total está limitado.
'El pr0 que le pasas al inicio es un valor inicial a partir del cúal corre el algoritmo.
'Nareas es el número de áreas para las que quieres hallar la pr0 a la que se igualan
Dim pr0old As Double
Dim i As Integer
Dim Ninitial As Double
Dim Nfin As Double
Dim proci As Double
Dim sumE As Double
Dim sumdE As Double
Dim tmp As Double
Dim i_area As Integer


For i = 0 To 99
    pr0old = pr0
    sumE = 0
    sumdE = 0
 
 'Extraer parámetros específicos de área y calcular esfuerzos, derivadas del esfuerzo, y sumatorios
   For i_area = 1 To Nfishedareas
      area = sortedfishedareas(i_area)
      Ninitial = BvulTmp(area)
      proci = pr0 + cost(area)
      tmp = (1 - handling * proci / price)
'       Nfin = proci / (price * q(Area) - q(Area) * handling * proci)
      Nfin = proci / (price * q(area) * tmp)
      sumE = sumE + handling * (Ninitial - Nfin) - Log(Nfin / Ninitial) / q(area)
      sumdE = sumdE + 1 / (q(area) * proci * tmp ^ 2)
   Next i_area
    'Nuevo valor de pr0 estimado
   pr0 = pr0old - (ET - sumE) / sumdE
   If (Abs(pr0 - pr0old) / pr0old < 0.00001) Then
        Exit For
   End If
Next i
calcPr0ET = pr0 'Devuelve valor de pr0 si converge el algoritmo y sale de la funcion
        
End Function

Function calcPr0TAC(pr0, sortedfishedareas, Nfishedareas, TACtmp)
'Función que estima la profitability a la que convergen las poblaciones en la Ideal Free Distribution
'sujeto a TAC.
'El pr0 que le pasas al inicio es un valor inicial a partir del cúal correr el algoritmo de N-R.
'Nfishedareas es el número de áreas para las que quieres hallar la pr0 a la que se igualan
Dim pr0old As Double
Dim i As Integer, i_area As Integer
Dim Nfin As Double
Dim proci As Double
Dim sumdf As Double
Dim sumcat As Double
Dim tmp As Double
   
   'next is Newtobn-Raphson iteration

For i = 0 To 99
   pr0old = pr0
   sumcat = 0
   sumdf = 0
   
   For i_area = 1 To Nfishedareas
     area = sortedfishedareas(i_area)
     proci = pr0 + cost(area)
     tmp = (1 - handling * proci / price)
     Nfin = proci / (price * q(area) * tmp)
     sumcat = sumcat + BvulTmp(area) - Nfin
   ' df = 1 / (q(Area) * price * (1 - handling * proci / price) ^ 2)
     sumdf = sumdf + Nfin / (proci * tmp)
   Next i_area
 'Nuevo valor de pr0 estimado
    pr0 = pr0old - (TACtmp - sumcat) / sumdf
 
    If (Abs(pr0 - pr0old) / pr0old < 0.00001) Then
        Exit For   'finish N-R iterations
    End If
Next i
calcPr0TAC = pr0

End Function
Function calcPr0TACandET(pr0, sortedfishedareas, Nfishedareas, TACtmp, ET, boundbyTAC)
'estima la profitability a la que convergen las poblaciones en la Ideal Free Distribution
'sujeto a TAC y a un tope de esfuerzo.
'El pr0 que le pasas al inicio es un valor inicial a partir del cúal correr el algoritmo de N-R.
'Nfishedareas es el número de áreas para las que quieres hallar la pr0 a la que se igualan

Dim pr0old As Double
Dim i As Integer, i_area As Integer
Dim Ninitial As Double
Dim Nfin As Double
Dim proci As Double
Dim sumdf As Double
Dim sumcat As Double
Dim sumE As Double
Dim sumdE As Double
Dim tmp As Double
Dim proE As Double

   
   'next is Newtobn-Raphson iteration

For i = 0 To 99
   pr0old = pr0
   sumcat = 0
   sumE = 0
   sumdf = 0
   sumdE = 0
   For i_area = 1 To Nfishedareas
     area = sortedfishedareas(i_area)
     proci = pr0 + cost(area)
     tmp = (1 - handling * proci / price)
     Nfin = proci / (price * q(area) * tmp)
     Ninitial = BvulTmp(area)
     sumcat = sumcat + Ninitial - Nfin
   ' df = 1 / (q(Area) * price * (1 - handling * proci / price) ^ 2)
     sumdf = sumdf + Nfin / (proci * tmp)
     sumE = sumE + handling * (Ninitial - Nfin) - Log(Nfin / Ninitial) / q(area)
     sumdE = sumdE + 1 / (q(area) * proci * tmp ^ 2)
   
   Next i_area
 
 'Nuevo valor de pr0 estimado
   
   pr0 = pr0old - (TACtmp - sumcat) / sumdf
   boundbyTAC = True
   
   If (sumE > ET) Then
      proE = pr0old - (ET - sumE) / sumdE
      If proE > pr0 Then
         pr0 = proE
         boundbyTAC = False
      End If
   End If
    
    If (Abs(pr0 - pr0old) / pr0old < 0.00001) Then
        Exit For   'finish N-R iterations
    End If
Next i
calcPr0TACandET = pr0

End Function

Function getE(pr0, area)
  
  Dim E As Double
  Dim N As Double
  
    N = getN(pr0, area)
    
    E = handling * (BvulTmp(area) - N) - Log(N / BvulTmp(area)) / q(area)
    If E < 0 Then 'Si el esfuerzo es negativo pasa a 0
        E = 0
    End If
    getE = E
End Function

Function getN(pr0, area)
 Dim N As Double
 Dim proci As Double

    proci = pr0 + cost(area)
        N = proci / (price * q(area) - q(area) * handling * proci)
    getN = N
End Function
Sub calcAll(pr0, area, year)
 Dim proci As Double
 Dim Ninitial As Double
 Dim rr As Integer
 
    rr = Region(area)
    proci = pr0 + cost(area)
    Ninitial = BvulTmp(area)
    BvulTmp(area) = proci / (price * q(area) - q(area) * handling * proci)
    HRTmp(area) = (1 - BvulTmp(area) / Ninitial)
    CatchTmp(area) = Ninitial - BvulTmp(area)
    EffortTmp(area) = handling * CatchTmp(area) - Log(BvulTmp(area) / Ninitial) / q(area)
    Catch(year, area) = Catch(year, area) + CatchTmp(area)
    AnnualCatch(rr) = AnnualCatch(rr) + CatchTmp(area)
    effort(year, area) = effort(year, area) + EffortTmp(area)
       
End Sub
