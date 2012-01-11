Attribute VB_Name = "M2_AnnualUpdate"
Sub Annual_update(year)
Attribute Annual_update.VB_ProcData.VB_Invoke_Func = " \n14"
Dim Area As Integer, age As Integer, i As Integer, rr As Integer, TEMP As Double

For rr = 1 To Nregions
    AnnualCatch(rr) = 0
Next rr

For Area = 1 To Nareas
    Bvulnerable(year + 1, Area) = BvulTmp(Area)
    Btotal(year + 1, Area) = BtotTmp(Area)
    
            
            ''''''''''''''''''''''' PROGRESSION DE COHORTES
     For age = Stage To AgePlus - 1
  '' Debug.Print muTmp(Area, age)
            n(year + 1, Area, age + 1) = NTmp(Area, age)
            mu(year + 1, Area, age + 1) = muTmp(Area, age)
            sd(year + 1, Area, age + 1) = sdTmp(Area, age)
            w(year + 1, Area, age + 1) = WTmp(Area, age)
    Next age
           
    For age = AgePlus - 1 To Stage Step -1
           '''''''''''''''''''''''PROGRESION DE HISTORIA DE PULSOS DE CADA COHORTE
            For i = 1 To Nfracs(Area, age)
               frac(Area, age + 1, i) = frac(Area, age, i)
                Z(Area, age + 1, i) = Z(Area, age, i)
            Next i
            Nfracs(Area, age + 1) = Nfracs(Area, age)
            
            For ilen = 1 To Nilens
                 pLage(Area, age + 1, ilen) = pLage(Area, age, ilen)
            Next ilen
            FracSel(Area, age + 1) = FracSel(Area, age)
     Next age
     '''''''''''''''''''''''fill in StAge
            Nfracs(Area, Stage) = 0
            mu(year + 1, Area, Stage) = mu(StYear, Area, Stage)
            sd(year + 1, Area, Stage) = sd(StYear, Area, Stage)
            w(year + 1, Area, Stage) = w(StYear, Area, Stage)
           For ilen = 1 To Nilens
                 pLage(Area, Stage, ilen) = pLStAge(Area, ilen)
            Next ilen
            FracSel(Area, Stage) = FracSelStAge(Area)
                    
        TEMP = n(year + 1, Area, AgePlus) + NTmp(Area, AgePlus)
        'mu(year + 1, area, AgePlus) = (mu(year + 1, area, AgePlus) * N(year + 1, area, AgePlus) + muTmp(area, AgePlus) * NTmp(area, AgePlus)) / TEMP
        
        For ilen = 1 To Nilens
            pLageplus(Area, ilen) = (pLage(Area, AgePlus, ilen) * n(year + 1, Area, AgePlus) + pLageplus(Area, ilen) * NTmp(Area, AgePlus)) / TEMP
        Next ilen
        
        n(year + 1, Area, AgePlus) = TEMP
          
   'rewrite mu,  W and FracSel for Ageplus with new size comp
            FracSel(Area, AgePlus) = 0
            For ilen = iLfull(Area) To Nilens
                FracSel(Area, AgePlus) = FracSel(Area, AgePlus) + pLageplus(Area, ilen)
            Next ilen
            mu(year + 1, Area, AgePlus) = 0
            w(year + 1, Area, AgePlus) = 0
            For ilen = 1 To Nilens
                    mu(year + 1, Area, AgePlus) = mu(year + 1, Area, AgePlus) + l(ilen) * pLageplus(Area, ilen)
                    w(year + 1, Area, AgePlus) = w(year + 1, Area, AgePlus) + W_L(Area, ilen) * pLageplus(Area, ilen)
                '   Debug.Print pLageplus(area, ilen)
            Next ilen
   
        For age = Stage To AgePlus
            NTmp(Area, age) = n(year + 1, Area, age)
            muTmp(Area, age) = mu(year + 1, Area, age)
            sdTmp(Area, age) = sd(year + 1, Area, age)
        Next age
    
   'compute biomasses excluding StAge (will be added later in calc_Rec in next year loop)
        'Bmature(year + 1, Area) = 0
        Btotal(year + 1, Area) = 0
        For age = Stage + 1 To AgePlus
            'Bmature(year + 1, Area) = Bmature(year + 1, Area) + n(year + 1, Area, age) * w(year + 1, Area, age) * FracMat(age)
            Btotal(year + 1, Area) = Btotal(year + 1, Area) + n(year + 1, Area, age) * w(year + 1, Area, age)
        Next age
       
Next Area
   For rr = 1 To Nregions
      ClosedRegionTmp(rr) = ClosedRegion(rr)
   Next rr
End Sub
Sub pLgen(year)
Attribute pLgen.VB_ProcData.VB_Invoke_Func = " \n14"
'Calculate size frequency distribution pL by area - only for output

For Area = 1 To Nareas
    intfact = 0
    For ilen = 1 To Nilens
        pL(year, Area, ilen) = 0
        For age = Stage To AgePlus
            pL(year, Area, ilen) = pL(year, Area, ilen) + pLage(Area, age, ilen) * NTmp(Area, age)
        Next age
        intfact = intfact + pL(year, Area, ilen)
    Next ilen
    
    For ilen = 1 To Nilens
            pL(year, Area, ilen) = pL(year, Area, ilen) / intfact
     '   Debug.Print pL(year, Area, ilen)
    Next ilen

Next Area

End Sub

