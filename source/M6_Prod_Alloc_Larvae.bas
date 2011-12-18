Attribute VB_Name = "M6_Prod_Alloc_Larvae"
Dim Area As Integer, i As Integer, Z() As Double, R0() As Double, B0() As Double

Sub Prod_Alloc_Larvae(year, SB)
'Toma biomasa desovante, produce larvas y las reparte
'puede ser relacion simplemente lineal
'
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

