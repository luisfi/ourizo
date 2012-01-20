Attribute VB_Name = "M2_AnnualUpdate"
Sub Annual_update(year)
Attribute Annual_update.VB_ProcData.VB_Invoke_Func = " \n14"
Dim area As Integer, age As Integer, i As Integer, rr As Integer, TEMP As Double

For rr = 1 To Nregions
    AnnualCatch(rr) = 0
Next rr

For area = 1 To Nareas
    Bvulnerable(year + 1, area) = BvulTmp(area)
    Btotal(year + 1, area) = BtotTmp(area)
    
            
            ''''''''''''''''''''''' PROGRESSION DE COHORTES
     For age = Stage To AgePlus - 1
  '' Debug.Print muTmp(Area, age)
            N(year + 1, area, age + 1) = NTmp(area, age)
            mu(year + 1, area, age + 1) = muTmp(area, age)
            sd(year + 1, area, age + 1) = sdTmp(area, age)
            w(year + 1, area, age + 1) = WTmp(area, age)
    Next age
           
    For age = AgePlus - 1 To Stage Step -1
           '''''''''''''''''''''''PROGRESION DE HISTORIA DE PULSOS DE CADA COHORTE
            For i = 1 To Nfracs(area, age)
               frac(area, age + 1, i) = frac(area, age, i)
                Z(area, age + 1, i) = Z(area, age, i)
            Next i
            Nfracs(area, age + 1) = Nfracs(area, age)
            
            For ilen = 1 To Nilens
                 pLage(area, age + 1, ilen) = pLage(area, age, ilen)
            Next ilen
            FracSel(area, age + 1) = FracSel(area, age)
     Next age
     '''''''''''''''''''''''fill in StAge
            Nfracs(area, Stage) = 0
            mu(year + 1, area, Stage) = mu(StYear, area, Stage)
            sd(year + 1, area, Stage) = sd(StYear, area, Stage)
            w(year + 1, area, Stage) = w(StYear, area, Stage)
           For ilen = 1 To Nilens
                 pLage(area, Stage, ilen) = pLStAge(area, ilen)
            Next ilen
            FracSel(area, Stage) = FracSelStAge(area)
                    
        TEMP = N(year + 1, area, AgePlus) + NTmp(area, AgePlus)
        'mu(year + 1, area, AgePlus) = (mu(year + 1, area, AgePlus) * N(year + 1, area, AgePlus) + muTmp(area, AgePlus) * NTmp(area, AgePlus)) / TEMP
        
        For ilen = 1 To Nilens
            pLageplus(area, ilen) = (pLage(area, AgePlus, ilen) * N(year + 1, area, AgePlus) + pLageplus(area, ilen) * NTmp(area, AgePlus)) / TEMP
        Next ilen
        
        N(year + 1, area, AgePlus) = TEMP
          
   'rewrite mu,  W and FracSel for Ageplus with new size comp
            FracSel(area, AgePlus) = 0
            For ilen = iLfull(area) To Nilens
                FracSel(area, AgePlus) = FracSel(area, AgePlus) + pLageplus(area, ilen)
            Next ilen
            mu(year + 1, area, AgePlus) = 0
            w(year + 1, area, AgePlus) = 0
            For ilen = 1 To Nilens
                    mu(year + 1, area, AgePlus) = mu(year + 1, area, AgePlus) + l(ilen) * pLageplus(area, ilen)
                    w(year + 1, area, AgePlus) = w(year + 1, area, AgePlus) + W_L(area, ilen) * pLageplus(area, ilen)
                '   Debug.Print pLageplus(area, ilen)
            Next ilen
   
        For age = Stage To AgePlus
            NTmp(area, age) = N(year + 1, area, age)
            muTmp(area, age) = mu(year + 1, area, age)
            sdTmp(area, age) = sd(year + 1, area, age)
        Next age
    
   'compute biomasses excluding StAge (will be added later in calc_Rec in next year loop)

        Btotal(year + 1, area) = 0
        For age = Stage + 1 To AgePlus
            Btotal(year + 1, area) = Btotal(year + 1, area) + N(year + 1, area, age) * w(year + 1, area, age)
        Next age
       
Next area
   For rr = 1 To Nregions
      ClosedRegionTmp(rr) = ClosedRegion(rr)
   Next rr
End Sub
Sub pLgen(year)
Attribute pLgen.VB_ProcData.VB_Invoke_Func = " \n14"
'Calculate size frequency distribution pL by area - only for output

For area = 1 To Nareas
    intfact = 0
    For ilen = 1 To Nilens
        pL(year, area, ilen) = 0
        For age = Stage To AgePlus
            pL(year, area, ilen) = pL(year, area, ilen) + pLage(area, age, ilen) * NTmp(area, age)
        Next age
        intfact = intfact + pL(year, area, ilen)
    Next ilen
    
    For ilen = 1 To Nilens
            pL(year, area, ilen) = pL(year, area, ilen) / intfact
     '   Debug.Print pL(year, Area, ilen)
    Next ilen

Next area

End Sub

