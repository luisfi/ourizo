Attribute VB_Name = "M6_Prod_Alloc_Larvae"
Dim area As Integer, i As Integer, Z() As Double, R0() As Double, B0() As Double

Sub Prod_Alloc_Larvae(year, Optional SB)
'Calcula biomasa desovante, produce larvas y las reparte
'puede ser relacion simplemente lineal
'
  If IsMissing(SB) Then 'función IsMissing() retorna el valor True si NO se ha enviado el parámetro que queremos comprobar
    For area = 1 To Nareas
        Bmature(year, area) = 0
        For age = Stage To AgePlus
            Bmature(year, area) = Bmature(year, area) + NTmp(area, age) * (1 - HRTmp(area) * FracSel(area, age) * FracHRPreRepr) * WTmp(area, age) * FracMat(age)
        Next age
    Next area
    SB = Bmature
  End If
   
  For area = 1 To Nareas
       Larvae(year, area) = SB(year, area) * ProdXB
  Next area
   
  For area = 1 To Nareas
    
       Settlers(year + Stage, area) = 0
    
       For i = 1 To Nareas
           Settlers(year + Stage, area) = Settlers(year + Stage, area) + Connect(area, i) * Larvae(year, i)
       Next i
  Next area

End Sub


