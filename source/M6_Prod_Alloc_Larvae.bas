Attribute VB_Name = "M6_Prod_Alloc_Larvae"
Dim Area As Integer, i As Integer, Z() As Double, R0() As Double, B0() As Double

Sub Prod_Alloc_Larvae(year, Optional SB)
'Calcula biomasa desovante, produce larvas y las reparte
'puede ser relacion simplemente lineal
'
  If IsMissing(SB) Then 'función IsMissing() retorna el valor True si NO se ha enviado el parámetro que queremos comprobar
    For Area = 1 To Nareas
        Bmature(year, Area) = 0
        For age = Stage To AgePlus
            Bmature(year, Area) = Bmature(year, Area) + NTmp(Area, age) * (1 - HRTmp(Area) * FracSel(Area, age) * FracHRPreRepr) * WTmp(Area, age) * FracMat(age)
        Next age
    Next Area
    SB = Bmature
  End If
   
  For Area = 1 To Nareas
       Larvae(year, Area) = SB(year, Area) * ProdXB
  Next Area
   
  For Area = 1 To Nareas
    
       Settlers(year + Stage, Area) = 0
    
       For i = 1 To Nareas
           Settlers(year + Stage, Area) = Settlers(year + Stage, Area) + Connect(Area, i) * Larvae(year, i)
       Next i
  Next Area

End Sub


