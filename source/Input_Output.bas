Attribute VB_Name = "Input_Output"
Option Explicit

Sub Output_Initialize()
Attribute Output_Initialize.VB_ProcData.VB_Invoke_Func = " \n14"
   Dim ilen As Integer, y
   Dim area As Integer, year As Integer, age As Integer
   
   Call Graph.clean("Output")
   Sheets.Add.Name = "Output"

'Print titulos para el output
     Worksheets("Output").Rows(1).Columns(4) = "Year"
     Worksheets("Output").Rows(1).Columns(2) = "Area"
     Worksheets("Output").Rows(1).Columns(3) = "Region"
     Worksheets("Output").Rows(1).Columns(1) = "Monte"
     Worksheets("Output").Rows(1).Columns(5) = "Catch"
     Worksheets("Output").Rows(1).Columns(6) = "Effort"
     Worksheets("Output").Rows(1).Columns(7) = "Bvulnerable"
     Worksheets("Output").Rows(1).Columns(8) = "Bmature"
     Worksheets("Output").Rows(1).Columns(9) = "Larvae"
     Worksheets("Output").Rows(1).Columns(10) = "Density"
     Worksheets("Output").Rows(1).Columns(11) = "Btotal"
     Worksheets("Output").Rows(1).Columns(12) = "Settlers"
     Worksheets("Output").Rows(1).Columns(13) = "Depletion_Bvul"
     Worksheets("Output").Rows(1).Columns(14) = "Depletion_Bmature"
     Worksheets("Output").Rows(1).Columns(15) = "HR"
     Worksheets("Output").Rows(1).Columns(16) = "Recruits"
     
   Call Graph.clean("Rotation")
   If RunFlags.Hstrategy = 1 Then 'If Rotational Strategy.
       Sheets.Add.Name = "Rotation"
       'Print titulos para el output
         Worksheets("Rotation").Rows(1).Columns(1) = "Monte"
         Worksheets("Rotation").Rows(1).Columns(2) = "Year"
         Worksheets("Rotation").Rows(1).Columns(3) = "Area"
         Worksheets("Rotation").Rows(1).Columns(4) = "Region"
         Worksheets("Rotation").Rows(1).Columns(5) = "Resting_Time"
         Worksheets("Rotation").Rows(1).Columns(6) = "Rotation_Period"
         Worksheets("Rotation").Rows(1).Columns(7) = "Adaptative_Period"
         For i = 1 To NOpenConditions
            Worksheets("Rotation").Rows(1).Columns(7 + i) = "ReOpenConditionValues_RC" & NOpenConditions
         Next i
         For i = 1 To Nreplicates
           For j = 1 To (Nyears * Nareas)
            Worksheets("Rotation").Rows(1 + (i - 1) * Nyears * Nareas + j).Columns(1) = i
           Next j
         Next i
    End If

   Call Graph.clean("Sizes")
   Sheets.Add.Name = "Sizes"

     For ilen = 1 To Nilens
        Worksheets("Sizes").Rows(1).Columns(10 + ilen) = l(ilen)
     Next ilen
  
   Call Graph.clean("Out_NAge_NSize")
   Sheets.Add.Name = "Out_NAge_NSize"
     Worksheets("Out_NAge_NSize").Rows(1).Columns(4) = "Year"
     Worksheets("Out_NAge_NSize").Rows(1).Columns(2) = "Area"
     Worksheets("Out_NAge_NSize").Rows(1).Columns(3) = "Region"
     Worksheets("Out_NAge_NSize").Rows(1).Columns(1) = "Monte"
            
    'Print indeces for ages
     For age = Stage To AgePlus
        Worksheets("Out_NAge_NSize").Rows(1).Columns(age + 4) = "Age " & age
     Next age
    'Print labels for size classes
     For ilen = 1 To Nilens
        Worksheets("Out_NAge_NSize").Rows(1).Columns(AgePlus + ilen + 4) = "Size" & l(ilen)
     Next ilen
   
   Call Graph.clean("mu_W")
   Sheets.Add.Name = "mu_W"

     'Print labels
     Worksheets("mu_W").Rows(1).Columns(1) = "Year"
     Worksheets("mu_W").Rows(1).Columns(2) = "Area"

     For age = Stage To AgePlus
       Worksheets("mu_W").Rows(1).Columns(age + 2) = "mu " & age
       Worksheets("mu_W").Rows(1).Columns(AgePlus + age + 3) = "W " & age
     Next age
     For year = 1 To Nyears
       For area = 1 To Nareas
         Worksheets("mu_W").Rows(year + 1 + (Nyears) * (area - 1)).Columns(1) = StYear - 1 + year
         Worksheets("mu_W").Rows(year + 1 + (Nyears) * (area - 1)).Columns(2) = area
       Next area
     Next year


End Sub

Sub Print_Initial_Conditions(FileName As String)

  Dim area As Integer, age As Integer, i As Integer
  
  Call Graph.clean(FileName)
  Sheets.Add.Name = FileName
     
     Worksheets(FileName).Rows(1).Columns(1) = "Area"
 
     For age = Stage To AgePlus
        Worksheets(FileName).Rows(1).Columns(age + 2) = "Age " & age
        Worksheets(FileName).Rows(1).Columns(AgePlus + age + 2) = "mu " & age
        Worksheets(FileName).Rows(1).Columns(2 * AgePlus + age + 2) = "W " & age
     Next age
     Worksheets(FileName).Rows(1).Columns(3 * AgePlus + 3) = "Btotal"
     Worksheets(FileName).Rows(1).Columns(3 * AgePlus + 4) = "Bmature"
     Worksheets(FileName).Rows(1).Columns(3 * AgePlus + 5) = "Bvulnerable"
     For i = 1 To Stage
        Worksheets(FileName).Rows(1).Columns(3 * AgePlus + 5 + i) = "Settlers(" & CStr(StYear + i - 1) & ")"
     Next i
 
   
       For area = 1 To Nareas
          For age = Stage To AgePlus
              Worksheets(FileName).Rows(area + 1).Columns(1) = area
              Worksheets(FileName).Rows(area + 1).Columns(age + 2) = N(StYear, area, age)
              Worksheets(FileName).Rows(area + 1).Columns(AgePlus + age + 2) = mu(StYear, area, age)
              Worksheets(FileName).Rows(area + 1).Columns(2 * AgePlus + age + 2) = w(StYear, area, age)
          Next age
          Worksheets(FileName).Rows(area + 1).Columns(3 * AgePlus + 3) = Btotal(StYear, area)
          Worksheets(FileName).Rows(area + 1).Columns(3 * AgePlus + 4) = Bmature(StYear, area)
          Worksheets(FileName).Rows(area + 1).Columns(3 * AgePlus + 5) = Bvulnerable(StYear, area)
         
         For i = 1 To Stage
            Worksheets(FileName).Rows(area + 1).Columns(3 * AgePlus + 5 + i) = Settlers(StYear + i - 1, area)
         Next i
       Next area
    
    Worksheets(FileName).Rows(Nareas + 2).Columns(1) = "NB!!!: Printted biomasses do not have contribution of StAge"
 
End Sub

Sub Read_Initial_Conditions(FileName As String)
   
   Dim area As Integer, age As Integer, i As Integer

    For area = 1 To Nareas
           For age = Stage To AgePlus
        
               N(StYear, area, age) = Worksheets(FileName).Rows(area + 1).Columns(age + 2)
               mu(StYear, area, age) = Worksheets(FileName).Rows(area + 1).Columns(AgePlus + age + 2)
               w(StYear, area, age) = Worksheets(FileName).Rows(area + 1).Columns(2 * AgePlus + age + 2)
    
           Next age
        
           Btotal(StYear, area) = Worksheets(FileName).Rows(area + 1).Columns(3 * AgePlus + 3)
           Bmature(StYear, area) = Worksheets(FileName).Rows(area + 1).Columns(3 * AgePlus + 4)
           Bvulnerable(StYear, area) = Worksheets(FileName).Rows(area + 1).Columns(3 * AgePlus + 5)
           For i = 1 To Stage
              Settlers(StYear + i - 1, area) = Worksheets(FileName).Rows(area + 1).Columns(3 * AgePlus + 5 + i)
           Next i
      Next area

End Sub

Sub Print_Output(monte)
Dim area As Integer, year As Integer, age As Integer
Dim TotBvulnerable As Double, TotBmature As Double, TotBtotal As Double, TotCatch As Double, ilen As Integer
Dim Density() As Double
ReDim Density(StYear To EndYear, Nareas) As Double
Dim mainpath As String
Application.ScreenUpdating = False

For year = 1 To Nyears
    For area = 1 To Nareas
        Worksheets("Output").Rows(year + 1 + (Nyears) * (area - 1) + (monte - 1) * Nyears * Nareas).Columns(7) = Bvulnerable(StYear - 1 + year, area)
        Worksheets("Output").Rows(year + 1 + (Nyears) * (area - 1) + (monte - 1) * Nyears * Nareas).Columns(8) = Bmature(StYear - 1 + year, area)
        Worksheets("Output").Rows(year + 1 + (Nyears) * (area - 1) + (monte - 1) * Nyears * Nareas).Columns(5) = Catch(StYear - 1 + year, area)
        Worksheets("Output").Rows(year + 1 + (Nyears) * (area - 1) + (monte - 1) * Nyears * Nareas).Columns(6) = effort(StYear - 1 + year, area)
        Worksheets("Output").Rows(year + 1 + (Nyears) * (area - 1) + (monte - 1) * Nyears * Nareas).Columns(9) = Larvae(StYear - 1 + year, area)
        
        For age = 1 To Nages
            Density(StYear - 1 + year, area) = Density(StYear - 1 + year, area) + N(StYear - 1 + year, area, age)
        Next age
        
        Density(StYear - 1 + year, area) = Density(StYear - 1 + year, area) / Surface(area)
        
        Worksheets("Output").Rows(year + 1 + (Nyears) * (area - 1) + (monte - 1) * Nyears * Nareas).Columns(10) = Density(StYear - 1 + year, area)
        Worksheets("Output").Rows(year + 1 + (Nyears) * (area - 1) + (monte - 1) * Nyears * Nareas).Columns(11) = Btotal(StYear - 1 + year, area)
        Worksheets("Output").Rows(year + 1 + (Nyears) * (area - 1) + (monte - 1) * Nyears * Nareas).Columns(12) = Settlers(StYear - 1 + year, area)
        Worksheets("Output").Rows(year + 1 + (Nyears) * (area - 1) + (monte - 1) * Nyears * Nareas).Columns(13) = Bvulnerable(StYear - 1 + year, area) / VBvirgin(area)
        Worksheets("Output").Rows(year + 1 + (Nyears) * (area - 1) + (monte - 1) * Nyears * Nareas).Columns(14) = Bmature(StYear - 1 + year, area) / SBvirgin(area)
        Worksheets("Output").Rows(year + 1 + (Nyears) * (area - 1) + (monte - 1) * Nyears * Nareas).Columns(15) = Catch(StYear - 1 + year, area) / Bvulnerable(StYear - 1 + year, area)
        Worksheets("Output").Rows(year + 1 + (Nyears) * (area - 1) + (monte - 1) * Nyears * Nareas).Columns(16) = N(StYear - 1 + year, area, Stage)
        
        Worksheets("Output").Rows(year + 1 + (Nyears) * (area - 1) + (monte - 1) * Nyears * Nareas).Columns(4) = StYear - 1 + year
        Worksheets("Output").Rows(year + 1 + (Nyears) * (area - 1) + (monte - 1) * Nyears * Nareas).Columns(2) = area
        Worksheets("Output").Rows(year + 1 + (Nyears) * (area - 1) + (monte - 1) * Nyears * Nareas).Columns(3) = Region(area)
        Worksheets("Output").Rows(year + 1 + (Nyears) * (area - 1) + (monte - 1) * Nyears * Nareas).Columns(1) = monte
        'Calculate totals accross areas per year
                
    Next area
Next year

If monte = Nreplicates Then

'START SAVING OUTPUT FOR SIMULATIONS
    Version = "Metapesca"
    Worksheets("Output").Activate
    Range(Cells(1, 1), Cells(Nreplicates * Nareas * Nyears + 1, 3 + AgePlus + 13 + Nilens)).Select
    Selection.Copy
    Workbooks.Add
    
    ActiveSheet.Paste
    
    Application.CutCopyMode = False
    
    mainpath = Workbooks(Version & ".xls").Path
    
    ChDir mainpath
    ActiveWorkbook.SaveAs FileName:=mainpath & "\" & "SimOut" & "\" & "Output" & ".csv", FileFormat:= _
        xlCSV, CreateBackup:=False
    ActiveWorkbook.Close
    Windows("Metapesca" & ".xls").Activate

'END SAVING OUTPUT FOR SIMULATIONS

Else
End If

Application.ScreenUpdating = True
End Sub

Sub Print_Output_NAge_NSize(monte)
Attribute Print_Output_NAge_NSize.VB_ProcData.VB_Invoke_Func = " \n14"
Dim area As Integer, year As Integer, age As Integer
Dim TotBvulnerable As Double, TotBmature As Double, TotBtotal As Double, TotCatch As Double, ilen As Integer
Dim Density() As Double
ReDim Density(StYear To EndYear, Nareas) As Double
Dim mainpath As String

Application.ScreenUpdating = False

      'Print 3d of N
     For year = 1 To Nyears
        For area = 1 To Nareas
           For age = Stage To AgePlus
             Worksheets("Out_NAge_NSize").Rows(year + 1 + (Nyears) * (area - 1) + (monte - 1) * Nyears * Nareas).Columns(age + 4) = N(StYear - 1 + year, area, age)
           Next age
        Next area
     Next year

     

For year = 1 To Nyears
    For area = 1 To Nareas
        
        For ilen = 1 To Nilens
            Worksheets("Out_NAge_NSize").Rows(year + 1 + (Nyears) * (area - 1) + (monte - 1) * Nyears * Nareas).Columns(AgePlus + ilen + 4) = pL(StYear - 1 + year, area, ilen)
        Next ilen
        
        Worksheets("Out_NAge_NSize").Rows(year + 1 + (Nyears) * (area - 1) + (monte - 1) * Nyears * Nareas).Columns(4) = StYear - 1 + year
        Worksheets("Out_NAge_NSize").Rows(year + 1 + (Nyears) * (area - 1) + (monte - 1) * Nyears * Nareas).Columns(2) = area
        Worksheets("Out_NAge_NSize").Rows(year + 1 + (Nyears) * (area - 1) + (monte - 1) * Nyears * Nareas).Columns(3) = Region(area)
        Worksheets("Out_NAge_NSize").Rows(year + 1 + (Nyears) * (area - 1) + (monte - 1) * Nyears * Nareas).Columns(1) = monte
        'Calculate totals accross areas per year
        
        '>>>>>>>>>>Commented out JLV 1/12/07
        'TotBvulnerable = Bvulnerable(StYear - 1 + year, Area) + TotBvulnerable
        'TotBmature = Bmature(StYear - 1 + year, Area) + TotBmature
        'TotBtotal = Btotal(StYear - 1 + year, Area) + TotBtotal
        'TotCatch = Catch(StYear - 1 + year, Area) + TotCatch
        '>>>>>>>>>>Commented out JLV 1/12/07
        
    Next area
Next year

If monte = Nreplicates Then

'START SAVING OUTPUT FOR SIMULATIONS
    Version = "Metapesca"
    Worksheets("Out_NAge_NSize").Activate
    Range(Cells(1, 1), Cells(Nreplicates * Nareas * Nyears + 1, 3 + AgePlus + 13 + Nilens)).Select
    Selection.Copy
    Workbooks.Add
    
    ActiveSheet.Paste
    
    Application.CutCopyMode = False
    
    mainpath = Workbooks(Version & ".xls").Path
    
    ChDir mainpath
    ActiveWorkbook.SaveAs FileName:=mainpath & "\" & "SimOut" & "\" & "Output_NAge_NSize" & ".csv", FileFormat:= _
        xlCSV, CreateBackup:=False
    ActiveWorkbook.Close
    Windows("Metapesca" & ".xls").Activate

Else
End If

'END SAVING OUTPUT FOR SIMULATIONS

Application.ScreenUpdating = True
End Sub

Sub Print_Input()
Attribute Print_Input.VB_ProcData.VB_Invoke_Func = " \n14"
   
   Version = "Metapesca"
   
    Dim ee As String, mainpath As String
    ee = Worksheets("Input").Rows(1).Columns(2)
    
    mainpath = Workbooks(Version & ".xls").Path
    
    Workbooks.Add
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs FileName:=mainpath & "\temp.xls", FileFormat:= _
        xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False _
        , CreateBackup:=False
    Windows(Version & ".xls").Activate
    Sheets("Input").Select
    Sheets("Input").Copy Before:=Workbooks("temp.xls").Sheets(1)
    'Application.WindowState = xlMinimized
    ActiveWorkbook.SaveAs FileName:=mainpath & "\" & "SimOut" & "\" & ee & ".dat", FileFormat:= _
        xlText, CreateBackup:=False
    ActiveWorkbook.Close

    
End Sub
Sub Print_Size_W()
Attribute Print_Size_W.VB_ProcData.VB_Invoke_Func = " \n14"

Dim area As Integer, year As Integer, age As Integer


'Print 3d of mu
For year = 1 To Nyears
    For area = 1 To Nareas
        For age = Stage To AgePlus
            Worksheets("mu_W").Rows(year + 1 + (Nyears) * (area - 1)).Columns(age + 2) = mu(StYear - 1 + year, area, age)
        Next age
    Next area
Next year
'Print 3d of W
For year = 1 To Nyears
    For area = 1 To Nareas
        For age = Stage To AgePlus
            Worksheets("mu_W").Rows(year + 1 + (Nyears) * (area - 1)).Columns(AgePlus + 1 + age + 2) = w(StYear - 1 + year, area, age)
        Next age
    Next area
Next year

For i = StYear To EndYear
    For ilen = 1 To Nilens
        Worksheets("Sizes").Rows(i - StYear + 2).Columns(10) = i
        Worksheets("Sizes").Rows(i - StYear + 2).Columns(10 + ilen) = pL(i, 1, ilen)
    Next ilen
Next i

End Sub

Sub Print_Rotational_Output(year)

  Dim area As Integer, LastRowColB As Integer
  
  Application.ScreenUpdating = False
  LastRowColB = Worksheets("Rotation").Range("B65536").End(xlUp).Row

  For area = 1 To Nareas

    Worksheets("Rotation").Rows(LastRowColB + area).Columns(2) = year
    Worksheets("Rotation").Rows(LastRowColB + area).Columns(3) = area
    Worksheets("Rotation").Rows(LastRowColB + area).Columns(4) = Region(area)
    Worksheets("Rotation").Rows(LastRowColB + area).Columns(5) = RestingTime(area)
    Worksheets("Rotation").Rows(LastRowColB + area).Columns(6) = RotationPeriod(area)
    Worksheets("Rotation").Rows(LastRowColB + area).Columns(7) = AdaptativeRotationFlag
      
    For i = 1 To NOpenConditions
        Worksheets("Rotation").Rows(LastRowColB + area).Columns(7 + i) = ReOpenConditionValues(i, area)
    Next i

  Next area
  Application.ScreenUpdating = True
End Sub
