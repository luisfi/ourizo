Attribute VB_Name = "M4_Calc_Recruits"
Dim Area As Integer, Recruits() As Double
Sub Deterministic_Recruits(year)
Attribute Deterministic_Recruits.VB_ProcData.VB_Invoke_Func = " \n14"
'Toma larvas que llegan a cada area y saca age 1 en cada area. Proceso local en cada area.
'Unica con un loop interno de area

ReDim Recruits(Nareas)
ReDim RecMax(Nareas) As Double

Select Case RunFlags.Rec

Case 1 ' Constant Recruitment

    For Area = 1 To Nareas
        N(year, Area, Stage) = R0(Area)
        NTmp(Area, Stage) = N(year, Area, Stage)
    Next Area

Case 2 'Compensation lineal

    For Area = 1 To Nareas
        RecMax(Area) = MinValue(MaxValue((Kcarga(Area) - Btotal(year, Area)) / w(StYear, Area, Stage), 0), Rmax(Area))
        Recruits(Area) = MinValue(RecMax(Area), Settlers(year, Area))
        N(year, Area, Stage) = Recruits(Area)
        NTmp(Area, Stage) = N(year, Area, Stage)
    Next Area

End Select

For Area = 1 To Nareas
'Add recruitment Biomass to totals
    Btotal(year, Area) = Btotal(year, Area) + N(year, Area, Stage) * w(year, Area, Stage)
    Bvulnerable(year, Area) = Bvulnerable(year, Area) + N(year, Area, Stage) * WvulStage(Area) * FracSel(Area, Stage)
    Bmature(year, Area) = Bmature(year, Area) + N(year, Area, Stage) * w(year, Area, Stage) * FracMat(Stage)
    BtotTmp(Area) = Btotal(year, Area)
    BvulTmp(Area) = Bvulnerable(year, Area)
Next Area

End Sub

Sub Random_Recruits(year)
Attribute Random_Recruits.VB_ProcData.VB_Invoke_Func = " \n14"
'Toma larvas que llegan a cada area y saca age 1 en cada area. Proceso local en cada area.
'Unica con un loop interno de area

ReDim Recruits(Nareas)
ReDim RecMax(Nareas) As Double

Select Case RunFlags.Rec

Case 1 ' Constant Recruitment

    For Area = 1 To Nareas
        N(year, Area, Stage) = R0(Area) * Exp(Rdev(year, Area) - 0.5 * RecCV ^ 2)
        NTmp(Area, Stage) = N(year, Area, Stage)
    Next Area

Case 2 'Compensation lineal

    For Area = 1 To Nareas
        RecMax(Area) = MinValue(MaxValue((Kcarga(Area) - Btotal(year, Area)) / w(StYear, Area, Stage), 0), Rmax(Area))
        Settlers(year, Area) = Settlers(year, Area) * Exp(Rdev(year, Area) - 0.5 * RecCV ^ 2)
        Recruits(Area) = MinValue(RecMax(Area), Settlers(year, Area))
        N(year, Area, Stage) = Recruits(Area)
        NTmp(Area, Stage) = N(year, Area, Stage)
    Next Area

End Select

For Area = 1 To Nareas
'Add recruitment Biomass to totals
    Btotal(year, Area) = Btotal(year, Area) + N(year, Area, Stage) * w(year, Area, Stage)
    Bvulnerable(year, Area) = Bvulnerable(year, Area) + N(year, Area, Stage) * WvulStage(Area) * FracSel(Area, Stage)
    Bmature(year, Area) = Bmature(year, Area) + N(year, Area, Stage) * w(year, Area, Stage) * FracMat(Stage)
    BtotTmp(Area) = Btotal(year, Area)
    BvulTmp(Area) = Bvulnerable(year, Area)
Next Area

End Sub
Sub Tunned_Recruits(year)
Attribute Tunned_Recruits.VB_ProcData.VB_Invoke_Func = " \n14"

'Usa reclutamientos tomados de un input file.
'It assumes observed recruitment without error

ReDim Recruits(Nareas)
ReDim RecMax(Nareas) As Double

'to escale recruitment time series

Select Case RunFlags.Rec

Case 1 ' Constant Recruitment

    For Area = 1 To Nareas
        Rdev(year, Area) = Log(ObsRec(year, Area) * q_Rec) - Log(R0(Area)) + 0.5 * RecCV ^ 2
        N(year, Area, Stage) = ObsRec(year, Area) * q_Rec
        NTmp(Area, Stage) = N(year, Area, Stage)
    Next Area

Case 2 'Compensation lineal

    For Area = 1 To Nareas
        RecMax(Area) = MinValue(MaxValue((Kcarga(Area) - Btotal(year, Area)) / w(StYear, Area, Stage), 0), Rmax(Area))
        Recruits(Area) = MinValue(RecMax(Area), Settlers(year, Area))
        Rdev(year, Area) = Log(ObsRec(year, Area) * q_Rec) - Log(Recruits(Area)) + 0.5 * RecCV ^ 2
        N(year, Area, Stage) = ObsRec(year, Area) * q_Rec
        NTmp(Area, Stage) = N(year, Area, Stage)
    Next Area

End Select

For Area = 1 To Nareas
'Add recruitment Biomass to totals
    Btotal(year, Area) = Btotal(year, Area) + N(year, Area, Stage) * w(year, Area, Stage)
    Bvulnerable(year, Area) = Bvulnerable(year, Area) + N(year, Area, Stage) * WvulStage(Area) * FracSel(Area, Stage)
    Bmature(year, Area) = Bmature(year, Area) + N(year, Area, Stage) * w(year, Area, Stage) * FracMat(Stage)
    BtotTmp(Area) = Btotal(year, Area)
    BvulTmp(Area) = Bvulnerable(year, Area)
Next Area

End Sub

Function MinValue(n1 As Double, n2 As Double) As Double
Attribute MinValue.VB_ProcData.VB_Invoke_Func = " \n14"

If n1 <= n2 Then
MinValue = n1
Else
MinValue = n2
End If

End Function

Function MaxValue(m1 As Double, m2 As Double) As Double
Attribute MaxValue.VB_ProcData.VB_Invoke_Func = " \n14"

If m1 >= m2 Then
MaxValue = m1
Else
MaxValue = m2
End If

End Function
