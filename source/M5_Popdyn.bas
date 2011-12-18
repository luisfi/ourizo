Attribute VB_Name = "M5_Popdyn"
Option Explicit
Option Base 0

'Declare indices
Dim age As Integer, year As Integer, Area As Integer, TEMP As Double, ilen As Integer, i As Integer, Lplus As Double, _
ilenplus As Integer
Dim XtraM() As Double, Wvul As Double, pLageplusTmp(), ZZ As Double, cumZZ As Double, sum_pLage As Double

Sub PopDyn(year)

ReDim XtraM(Nareas), pLageplusTmp(Nilens)
   
    For Area = 1 To Nareas
          
          XtraM(Area) = 0
          
        'Optar por tipo de crecimiento
        Select Case RunFlags.Growth_type
          Case 1
            'Density-independent growth
                g(Area) = 1
            
          Case 2
            'Lineal density dependence
                                
            If BtotTmp(Area) < Bthreshold(Area) Then
                g(Area) = 1
                
            ElseIf BtotTmp(Area) > (Bg0(Area)) Then
                g(Area) = 0
            Else
                g(Area) = 1 - ((1 - gk(Area)) / (Kcarga(Area) - Bthreshold(Area)) * (BtotTmp(Area) - Bthreshold(Area)))
            End If
                    
            'Debug.Print g(area)
            'Debug.Print Alpha(area), Beta(area)
                    
          End Select
        
            Alpha(Area) = (1 - Rho(Area)) * Linf(Area) * g(Area)
            Beta(Area) = 1 - (1 - Rho(Area)) * g(Area)
  
        
        For age = Stage To AgePlus - 1
            
            
            If Flag_Rec_Fish(Area, age) < 3 Then
                ZZ = (Lfull(Area) - muTmp(Area, age)) / sdTmp(Area, age)
                cumZZ = 1 - Cumd_Norm(ZZ)
            
                If (cumZZ > 0.02 And cumZZ < 0.98) Then
                        Flag_Rec_Fish(Area, age) = 2
                ElseIf cumZZ > 0.98 Then
                        Flag_Rec_Fish(Area, age) = 3
                End If
            
            Else
            End If
    
    If HRTmp(Area) > 0 Then
    
        If Flag_Rec_Fish(Area, age) = 2 Then
            For i = 1 To Nfracs(Area, age)
               frac(Area, age, i + 1) = frac(Area, age, i) * (1 - HRTmp(Area))
               Z(Area, age, i + 1) = Z(Area, age, i)
            Next i
            frac(Area, age, 1) = (1 - HRTmp(Area))
            Z(Area, age, 1) = ZZ
            Nfracs(Area, age) = Nfracs(Area, age) + 1
        End If
            
            
''Debug.Print muTmp(Area, age)
'''''''''''''''''''''''''''''''''''''''NOW THEY GROW AND DIE!
            NTmp(Area, age) = NTmp(Area, age) * Exp(-(M(Area))) * (1 - HRTmp(Area) * FracSel(Area, age))
                               
    Else 'En caso que HRTmp = 0 se simplifica a:
                   
            NTmp(Area, age) = NTmp(Area, age) * Exp(-(M(Area)))
                
    End If
                
                muTmp(Area, age) = Alpha(Area) + Beta(Area) * muTmp(Area, age)
                sdTmp(Area, age) = CVmu(Area) * muTmp(Area, age)
                                  
''Debug.Print muTmp(Area, age)
             Select Case Flag_Rec_Fish(Area, age)
                    Case 1   'Not recruited
                           Call M8_Library.Norm(Area, age)
                    Case 2 'Partially recruited
                            Call M8_Library.Trunc_Norm(Area, age)
                    Case 3   'Fully recruited
                            Call M8_Library.Trunc_Norm(Area, age)
             End Select
        
          Next age
         
     'Ageplus need to be dealt with separately because it has an explicit size-comp vector
     ' equal to pLAgeplus
             
        'size-selective fishing
           For ilen = iLfull(Area) To Nilens
                  pLageplus(Area, ilen) = (1 - HRTmp(Area)) * pLageplus(Area, ilen)
            Next ilen
                                       ' NB: no need to integrate until after growth takes place
       'growth
           For ilen = 1 To Nilens
                   pLageplusTmp(ilen) = 0
                ' Debug.Print pLageplus(area, ilen)
           Next ilen
                    
           ilen = 1
           While l(ilen) < Linf(Area) And ilen < Nilens
              Lplus = Alpha(Area) + Beta(Area) * (l(ilen) + 0 * Linc) 'l(ilen) is the center of interval
                ilenplus = 1 + (Lplus - L1) / Linc     'this rounds the number (doesn't truncate)
                If (ilenplus > Nilens) Then ilenplus = Nilens
                pLageplusTmp(ilenplus) = pLageplusTmp(ilenplus) + pLageplus(Area, ilen)
              ilen = ilen + 1
           Wend
            
           For i = ilen To Nilens       'NB! these are for l(ilen) >= Linf
                pLageplusTmp(i) = pLageplusTmp(i) + pLageplus(Area, i)
           Next i
            
         ' now normalize
            TEMP = 0
            For ilen = 1 To Nilens
                 TEMP = TEMP + pLageplusTmp(ilen)
            Next ilen
            
            For ilen = 1 To Nilens
            
                'Debug.Print pLageplus(area, ilen)
                             
                 pLageplus(Area, ilen) = pLageplusTmp(ilen) / TEMP
                 pLage(Area, AgePlus, ilen) = pLageplus(Area, ilen)
            
           ' Debug.Print pLageplus(area, ilen)
            
            Next ilen
                                   
                                   
       '     For ilen = 1 To Nilens
            
        '        Debug.Print pLageplus(area, ilen)
            
         '   Next ilen
                                   
                                   
                                                                      
                                   
            NTmp(Area, AgePlus) = NTmp(Area, AgePlus) * Exp(-(M(Area))) * (1 - HRTmp(Area) * FracSel(Area, AgePlus))
        '
        'Now compute FracSel, weights and biomasses for next year
         If Flag_Rec_Fish(Area, age) < 3 Then
            For age = Stage To AgePlus
               FracSel(Area, age) = 0
               For ilen = iLfull(Area) To Nilens
                   FracSel(Area, age) = FracSel(Area, age) + pLage(Area, age, ilen)
               Next ilen
            Next age
          End If
          
        BtotTmp(Area) = 0
        BvulTmp(Area) = 0
        
        For age = Stage To AgePlus
       
            WTmp(Area, age) = 0
            For ilen = 1 To Nilens
                WTmp(Area, age) = WTmp(Area, age) + W_L(Area, ilen) * pLage(Area, age, ilen)
            Next ilen
                        
            Wvul = 0
            For ilen = iLfull(Area) To Nilens
                Wvul = Wvul + W_L(Area, ilen) * pLage(Area, age, ilen)
        'Debug.Print pLage(area, age, ilen)
            Next ilen
        
            BtotTmp(Area) = BtotTmp(Area) + NTmp(Area, age) * WTmp(Area, age)
            BvulTmp(Area) = BvulTmp(Area) + NTmp(Area, age) * Wvul * FracSel(Area, age)
        'Debug.Print Wvul
        'Debug.Print FracSel(area,age)
        'Debug.Print BvulTmp(area)
        Next age
         
         'Evaluar si Btotal esta por encima de K
        
        Kcarga_adults(Area) = Kcarga(Area) - R0(Area) * w(StYear, Area, Stage)
        
        If (BtotTmp(Area) > Kcarga_adults(Area)) Then
            XtraM(Area) = Kcarga_adults(Area) / BtotTmp(Area)
            BtotTmp(Area) = BtotTmp(Area) * XtraM(Area)
            BvulTmp(Area) = BvulTmp(Area) * XtraM(Area)
            For age = Stage To AgePlus
                    NTmp(Area, age) = NTmp(Area, age) * XtraM(Area)
            Next age
        
       Else
       End If
       'Debug.Print BvulTmp(area)
    Next Area

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
