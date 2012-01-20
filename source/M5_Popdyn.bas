Attribute VB_Name = "M5_Popdyn"
Option Explicit
Option Base 0

'Declare indices
Dim age As Integer, year As Integer, area As Integer, TEMP As Double, ilen As Integer, i As Integer, Lplus As Double, _
ilenplus As Integer
Dim XtraM() As Double, Wvul As Double, pLageplusTmp(), ZZ As Double, cumZZ As Double, sum_pLage As Double

Sub PopDyn(year)

   ReDim XtraM(Nareas), pLageplusTmp(Nilens)
   
   For area = 1 To Nareas
          
      XtraM(area) = 0
          
      Select Case RunFlags.Growth_type
        Case 1
            'Density-independent growth
                g(area) = 1
            
        Case 2
            'Lineal density dependence
                                
            If BtotTmp(area) < Bthreshold(area) Then
                g(area) = 1
                
            ElseIf BtotTmp(area) > (Bg0(area)) Then
                g(area) = 0
            Else
                g(area) = 1 - ((1 - gk(area)) / (Kcarga(area) - Bthreshold(area)) * (BtotTmp(area) - Bthreshold(area)))
            End If
                    
            'Debug.Print g(area)
            'Debug.Print Alpha(area), Beta(area)
                    
      End Select
        
      Alpha(area) = (1 - Rho(area)) * Linf(area) * g(area)
      Beta(area) = 1 - (1 - Rho(area)) * g(area)
  
      For age = Stage To AgePlus - 1
             
         If flag_Partial_Rec(area, age) < 3 Then    'age not fully recruited
             ZZ = (Lfull(area) - muTmp(area, age)) / sdTmp(area, age)
             cumZZ = 1 - Cumd_Norm(ZZ)
            
             If (cumZZ > 0.02 And cumZZ < 0.98) Then
                   flag_Partial_Rec(area, age) = 2
             ElseIf cumZZ > 0.98 Then
                   flag_Partial_Rec(area, age) = 3
             End If
        
         End If
    
         If HRTmp(area) > 0 Then
    
             If flag_Partial_Rec(area, age) = 2 Then    'age partially recruited
                For i = 1 To Nfracs(area, age)
                   frac(area, age, i + 1) = frac(area, age, i) * (1 - HRTmp(area))
                   Z(area, age, i + 1) = Z(area, age, i)
                Next i
                frac(area, age, 1) = (1 - HRTmp(area))
                Z(area, age, 1) = ZZ
                Nfracs(area, age) = Nfracs(area, age) + 1
             End If
            
''''''NOW THEY GROW AND DIE
             NTmp(area, age) = NTmp(area, age) * Exp(-(M(area))) * (1 - HRTmp(area) * FracSel(area, age))
                               
         Else ' HRTmp = 0
                   
              NTmp(area, age) = NTmp(area, age) * Exp(-(M(area)))
                
         End If
           
         muTmp(area, age) = Alpha(area) + Beta(area) * muTmp(area, age)
         sdTmp(area, age) = CVmu(area) * muTmp(area, age)
                                  
''Debug.Print muTmp(Area, age)
         Select Case flag_Partial_Rec(area, age)
             Case 1   'Not recruited
                 Call M8_Library.Norm(area, age)
             Case 2 'Partially recruited
                 Call M8_Library.Trunc_Norm(area, age)
             Case 3 'Fully recruited
                 Call M8_Library.Trunc_Norm(area, age)
         End Select
        
      Next age
         
     'Ageplus needs to be dealt with separately because it has an explicit size-comp vector
     ' equal to pLAgeplus
             
      'size-selective fishing
      For ilen = iLfull(area) To Nilens
         pLageplus(area, ilen) = (1 - HRTmp(area)) * pLageplus(area, ilen)
      Next ilen
                                       ' NB: no need to integrate until after growth takes place
      'growth
      For ilen = 1 To Nilens
          pLageplusTmp(ilen) = 0
                ' Debug.Print pLageplus(area, ilen)
      Next ilen
                    
      ilen = 1
      While l(ilen) < Linf(area) And ilen < Nilens
         Lplus = Alpha(area) + Beta(area) * (l(ilen) + 0 * Linc) 'l(ilen) is the center of interval
         ilenplus = 1 + (Lplus - L1) / Linc     'this rounds the number (doesn't truncate)
         If (ilenplus > Nilens) Then ilenplus = Nilens
         pLageplusTmp(ilenplus) = pLageplusTmp(ilenplus) + pLageplus(area, ilen)
         ilen = ilen + 1
      Wend
            
      For i = ilen To Nilens       'NB! these are for l(ilen) >= Linf
          pLageplusTmp(i) = pLageplusTmp(i) + pLageplus(area, i)
      Next i
            
        ' now normalize
      TEMP = 0
      For ilen = 1 To Nilens
          TEMP = TEMP + pLageplusTmp(ilen)
      Next ilen
            
      For ilen = 1 To Nilens
                                
          pLageplus(area, ilen) = pLageplusTmp(ilen) / TEMP
          pLage(area, AgePlus, ilen) = pLageplus(area, ilen)
                      
      Next ilen
                                   
      NTmp(area, AgePlus) = NTmp(area, AgePlus) * Exp(-(M(area))) * (1 - HRTmp(area) * FracSel(area, AgePlus))
        
      'Now compute FracSel, weights and biomasses for next period
      If flag_Partial_Rec(area, age) < 3 Then
          For age = Stage To AgePlus
             FracSel(area, age) = 0
             For ilen = iLfull(area) To Nilens
                FracSel(area, age) = FracSel(area, age) + pLage(area, age, ilen)
             Next ilen
          Next age
      End If
          
      BtotTmp(area) = 0
      BvulTmp(area) = 0
        
      For age = Stage To AgePlus
       
         WTmp(area, age) = 0
         For ilen = 1 To Nilens
             WTmp(area, age) = WTmp(area, age) + W_L(area, ilen) * pLage(area, age, ilen)
         Next ilen
                        
         Wvul = 0
         For ilen = iLfull(area) To Nilens
             Wvul = Wvul + W_L(area, ilen) * pLage(area, age, ilen)
         Next ilen
        
         BtotTmp(area) = BtotTmp(area) + NTmp(area, age) * WTmp(area, age)
         BvulTmp(area) = BvulTmp(area) + NTmp(area, age) * Wvul * FracSel(area, age)
      
      Next age
         
      'apply extra mortality if Btotal exceeded carrying capacity
        
      Kcarga_adults(area) = Kcarga(area) - R0(area) * w(StYear, area, Stage)
        
      If (BtotTmp(area) > Kcarga_adults(area)) Then
          XtraM(area) = Kcarga_adults(area) / BtotTmp(area)
          BtotTmp(area) = BtotTmp(area) * XtraM(area)
          BvulTmp(area) = BvulTmp(area) * XtraM(area)
            
          For age = Stage To AgePlus
              NTmp(area, age) = NTmp(area, age) * XtraM(area)
          Next age
        
      End If
       'Debug.Print BvulTmp(area)
   Next area

End Sub

Sub Maturity(AgeFullMature, FracMat)
Attribute Maturity.VB_ProcData.VB_Invoke_Func = " \n14"

ReDim FracMat(Stage To AgePlus)
     
    For age = Stage To AgePlus
        If age >= AgeFullMature Then
            FracMat(age) = 1
        Else
            FracMat(age) = 0
        End If
    Next age

End Sub
