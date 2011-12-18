Attribute VB_Name = "Graph"
Option Explicit

Sub Graphs()
Attribute Graphs.VB_ProcData.VB_Invoke_Func = " \n14"
Application.ScreenUpdating = False
Application.DisplayAlerts = False
   Dim StRowTable As Integer, EndRowTable As Integer, StColTable As Integer, EndColTable As Integer
   
   Dim Total As Double, TotalCatch As Double, TotalVB As Double, TotalSB As Double
   
   Dim Area As Integer, RowsTable As Integer, NNombres As Integer
   Dim vertoffsets As Integer, horzoffsets As Integer
   Dim nameGraph As String, vertPos As Integer, horzPos As Integer
   Dim StarthorzRng As Integer, EndhorzRng As Integer
   Dim Xrange As String, Datarange As String, i As Integer, j As Integer, jj As Integer
   Dim theRange As Range, theChart As ChartObject
   Dim posver As Integer, poshorz As Integer
   Dim gheight As Integer, gwidth As Integer
   Dim etiquetas() As String
   ReDim etiquetas(Nareas + 1)
   Dim Region_mapa As String
    
    GraphFlag = Worksheets("Input").Rows(8).Columns(2)
    
    If GraphFlag = 2 Then GoTo 2

' Catch/Effort/VulnerableBiomass/SpawningBiomass/Larvae
' Go and get the dimensioning parameters
   
   
   vertoffsets = 2
   horzoffsets = 1
   Xrange = "B" & horzoffsets + 1 & ":B" & horzoffsets + Nyears
   
   Datarange = Xrange & "," & StarthorzRng & ":" & EndhorzRng
   
 gheight = 250
 gwidth = 250
  
  Call clean("Mapas")
  
  Call clean("Graphs")
  Sheets.Add.Name = "Graphs"
  Worksheets("Graphs").Activate
 'Delete Table

   With Worksheets("Output")
     On Error Resume Next
     Range("CU2:FX500").Delete
   End With
    
       
   For i = 1 To 11
    
    Dim Title As String, Xaxes As String, Yaxes As String
    
    Select Case i
      Case 1
        Title = "Catch"
        Yaxes = "Catch"
        Xaxes = "Time"
        posver = 1
        poshorz = 1
      
      Case 2
        Title = "Effort"
        Yaxes = "Effort"
        Xaxes = "Time"
        posver = 250
        poshorz = 1
 
      Case 3
        Title = "Vulnerable Biomass"
        Yaxes = "Vulnerable Biomass"
        Xaxes = "Time"
        posver = 500
        poshorz = 500
       
      Case 4
        Title = "Spawning Biomass"
        Yaxes = "Spawning Biomass"
        Xaxes = "Time"
        posver = 1
        poshorz = 250
     
      Case 5
        Title = "Larvae"
        Yaxes = "Larvae"
        Xaxes = "Time"
        posver = 250
        poshorz = 250
      
      Case 6
        Title = "Density"
        Yaxes = "Density"
        Xaxes = "Time"
        posver = 1
        poshorz = 500
      
      Case 7
        Title = "Recruits"
        Yaxes = "Recruits"
        Xaxes = "Time"
        posver = 500
        poshorz = 250
      
      Case 8
        Title = "Total Biomass"
        Yaxes = "Total Biomass"
        Xaxes = "Time"
        posver = 250
        poshorz = 500
      
      Case 9
        Title = "Harvest Rate"
        Yaxes = "Harvest Rate"
        Xaxes = "Time"
        posver = 500
        poshorz = 1
            
      Case 10
        Title = "Depletion Bvul"
        Yaxes = "Depletion Bvul"
        Xaxes = "Time"
        posver = 1
        poshorz = 750
            
      Case 11
        Title = "Depletion Bmat"
        Yaxes = "Depletion Bmat"
        Xaxes = "Time"
        posver = 250
        poshorz = 750
            
      Case Else
      
    End Select
 
 'Create Tables
 'write years
     
     For jj = 1 To Nyears
        Worksheets("Output").Cells(1 + jj + (i - 1) * Nyears, 99).Value = Worksheets("Output").Cells(1 + jj, 4).Value
     Next
          
     
       etiquetas(Nareas + 1) = "Total"
     
     For j = 1 To Nareas
 
        'Writes area index
       etiquetas(j) = Worksheets("Input").Cells(42, 1 + j).Value
       
       For jj = 1 To Nyears
         Select Case i
         Case 7
            vertoffsets = 16
            horzoffsets = 1 + (j - 1) * Nyears
            Worksheets("Output").Cells(1 + (i - 1) * Nyears + jj, 99 + j).Value = Worksheets("Output").Cells(horzoffsets + jj, vertoffsets).Value
         
         Case 8
            vertoffsets = 11
            horzoffsets = 1 + (j - 1) * Nyears
            Worksheets("Output").Cells(1 + (i - 1) * Nyears + jj, 99 + j).Value = Worksheets("Output").Cells(horzoffsets + jj, vertoffsets).Value
         
         Case 9
            vertoffsets = 15
            horzoffsets = 1 + (j - 1) * Nyears
            Worksheets("Output").Cells(1 + (i - 1) * Nyears + jj, 99 + j).Value = Worksheets("Output").Cells(horzoffsets + jj, vertoffsets).Value
         
         Case 10
            vertoffsets = 13
            horzoffsets = 1 + (j - 1) * Nyears
            Worksheets("Output").Cells(1 + (i - 1) * Nyears + jj, 99 + j).Value = Worksheets("Output").Cells(horzoffsets + jj, vertoffsets).Value
         
         Case 11
            vertoffsets = 14
            horzoffsets = 1 + (j - 1) * Nyears
            Worksheets("Output").Cells(1 + (i - 1) * Nyears + jj, 99 + j).Value = Worksheets("Output").Cells(horzoffsets + jj, vertoffsets).Value
                  
         Case Else
            vertoffsets = 5 + (i - 1)
            horzoffsets = 1 + (j - 1) * Nyears
            Worksheets("Output").Cells(1 + (i - 1) * Nyears + jj, 99 + j).Value = Worksheets("Output").Cells(horzoffsets + jj, vertoffsets).Value
         End Select
       Next
     Next
    
    'xxxxxxx<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>
    'Calculate Total (REVISAR CALCULA TOTAL ANTES DE QUE ESTEN LOS PARCIALES!!!!!!!!!)
    'xxxxxxx
    
    
    For RowsTable = 1 To 8 * Nyears
        Total = 0
        For Area = 1 To Nareas
            Total = Total + Worksheets("Output").Cells(1 + RowsTable, 99 + Area).Value
            Worksheets("Output").Cells(1 + RowsTable, 100 + Nareas).Value = Total
        Next
    Next
    
    For RowsTable = 8 * Nyears + 1 To 9 * Nyears
        TotalCatch = 0
        TotalVB = 0
        For Area = 1 To Nareas
            TotalCatch = TotalCatch + Worksheets("Output").Cells(1 + RowsTable - 8 * Nyears, 99 + Area).Value
            TotalVB = TotalVB + Worksheets("Output").Cells(1 + RowsTable - 6 * Nyears, 99 + Area).Value
        Next
        Worksheets("Output").Cells(1 + RowsTable, 100 + Nareas).Value = TotalCatch / TotalVB
    Next
    
    For RowsTable = 9 * Nyears + 1 To 10 * Nyears
        TotalVB = 0
        For Area = 1 To Nareas
            TotalVB = TotalVB + Worksheets("Output").Cells(1 + RowsTable - 7 * Nyears, 99 + Area).Value
        Next
        'Debug.Print TotalVB
        Worksheets("Output").Cells(1 + RowsTable, 100 + Nareas).Value = TotalVB / VB0_all
    Next
            
    For RowsTable = 10 * Nyears + 1 To 11 * Nyears
        TotalSB = 0
        For Area = 1 To Nareas
            TotalSB = TotalSB + Worksheets("Output").Cells(1 + RowsTable - 7 * Nyears, 99 + Area).Value
        Next
        Worksheets("Output").Cells(1 + RowsTable, 100 + Nareas).Value = TotalSB / SB0_all
    Next

     Select Case i
     Case 6
       StRowTable = 2 + Nyears * (i - 1)
       EndRowTable = 1 + Nyears * (i - 1) + Nyears
       StColTable = 99
       EndColTable = 99 + Nareas
     Case Else
        StRowTable = 2 + Nyears * (i - 1)
        EndRowTable = 1 + Nyears * (i - 1) + Nyears
        StColTable = 99
        EndColTable = 100 + Nareas
     End Select

    
    

            
  
If GraphFlag = 3 Then
    GoTo 3
ElseIf GraphFlag = 2 Then
    GoTo 2
End If

 'Create Graph
     Worksheets("Output").Activate
     Set theRange = Worksheets("Output").Range(Cells(StRowTable, StColTable), Cells(EndRowTable, EndColTable))
     Set theChart = Worksheets("Graphs").ChartObjects.Add(posver, poshorz, gheight, gwidth)
 
 'Format Graph
 
        theChart.Activate

        Worksheets("Graphs").Activate
        With theChart.Chart
         .ChartType = xlXYScatterLines
         .SetSourceData Source:=theRange, PlotBy _
             :=xlColumns

         .HasTitle = True
         .ChartTitle.Characters.Text = Title
         .Axes(xlCategory, xlPrimary).HasTitle = True
         .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = Xaxes
         .Axes(xlValue, xlPrimary).HasTitle = True
         .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = Yaxes
         End With

            theChart.Activate
            ActiveChart.ChartArea.Select
            ActiveChart.PlotArea.Select
            
'            If i = 6 Then
'                NNombres = Nareas
'            Else
'                NNombres = Nareas + 1
'            End If
            
            If i = 6 Then
                NNombres = Nareas
            Else
                NNombres = Nareas + 1
                
                    ActiveChart.SeriesCollection(6).Select
                    With Selection.Border
                        .ColorIndex = 57
                        .Weight = xlThick
                        .LineStyle = xlContinuous
                    End With
                    With Selection
                        .MarkerBackgroundColorIndex = xlAutomatic
                        .MarkerForegroundColorIndex = xlAutomatic
                        .MarkerStyle = xlNone
                        .Smooth = False
                        .MarkerSize = 9
                        .Shadow = False
                    End With
            End If
            
            For jj = 1 To NNombres
                ActiveChart.SeriesCollection(jj).Name = etiquetas(jj)
            Next
            
' Now lets format the graph
            
 
     With ActiveChart.Axes(xlCategory)
        .MinimumScale = StYear
        .MaximumScale = EndYear
        .MinorUnitIsAuto = False
        .MajorUnit = Int(Nyears / 5)
        .Crosses = xlAutomatic
        .ReversePlotOrder = False
        .ScaleType = xlLinear
        .DisplayUnit = xlNone
       
    End With
    
    ActiveChart.Axes(xlCategory).Select
    Selection.TickLabels.AutoScaleFont = True
    With Selection.TickLabels.Font
        .Name = "Arial"
        .FontStyle = "Regular"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .Background = xlAutomatic
    End With
    
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    Selection.Delete
    
    ActiveChart.Axes(xlValue).Select
    With ActiveChart.Axes(xlValue)
        .MinimumScale = 0
        .MaximumScaleIsAuto = True
        .MinorUnitIsAuto = True
        .MajorUnitIsAuto = True
        .Crosses = xlAutomatic
        .ReversePlotOrder = False
        .ScaleType = xlLinear
        .DisplayUnit = xlNone
    End With
    
    ActiveChart.PlotArea.Select
    With Selection.Border
        .ColorIndex = 16
        .Weight = xlThin
        .LineStyle = xlContinuous
    End With
    Selection.Interior.ColorIndex = xlNone

    ActiveChart.Legend.Select
    Selection.AutoScaleFont = True
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "Regular"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .Background = xlAutomatic
    End With
    Selection.Position = xlBottom
   
3

Next

Worksheets("Graphs").Rows(1).Columns(1).Select

If GraphFlag = 1 Then GoTo 2
'################################################################################################
'Creacion de grafico espacial

'Borrado de datos para grafico y hoja conteniendo el grafico

Dim VarEspacial As Integer
Dim EspTitle As String

VarEspacial = Worksheets("Input").Rows(8).Columns(3)
Nareas = Worksheets("Input").Rows(31).Columns(2)

Select Case VarEspacial

Case 1
    EspTitle = "Captura"
Case 2
    EspTitle = "Esfuerzo"
Case 3
    EspTitle = "Bvulnerable"
Case 4
    EspTitle = "Bmature"
Case 5
    EspTitle = "Larvas"
Case 6
    EspTitle = "Densidad"
Case 7
    EspTitle = "Reclutas"
Case 8
    EspTitle = "Btotal"
Case 9
    EspTitle = "HR"
Case 10
    EspTitle = "Depletion Bvul"
Case 11
    EspTitle = "Depletion Bmat"

End Select


Worksheets("Graphs").Rows(3).Columns(99) = Worksheets("Output").Rows(2).Columns(99)
For Area = 1 To Nareas

'Longitud
    Worksheets("Graphs").Rows(1).Columns(99 + Area) = Abs(Worksheets("Input").Rows(45).Columns(Area + 1))

'Latitud
    Worksheets("Graphs").Rows(2).Columns(99 + Area) = Worksheets("Input").Rows(44).Columns(Area + 1)

'Variable a graficar determinante del tamanio de las burbujas
    Worksheets("Graphs").Rows(3).Columns(99 + Area) = Sqr(Worksheets("Output").Rows(2 + ((VarEspacial - 1) * Nyears)).Columns(Area + 99))

Next

  Range("CV1:IV3").Select
    Charts.Add
    ActiveChart.ChartType = xlBubble
    ActiveChart.SetSourceData Source:=Sheets("Graphs").Range("CV1:IV3"), PlotBy:= _
        xlRows
    ActiveChart.SeriesCollection(1).Name = "=Graphs!R3C99"
    ActiveChart.Location Where:=xlLocationAsNewSheet, Name:="Mapas"
    
       
    With ActiveChart
        .HasTitle = True
        .ChartTitle.Characters.Text = EspTitle
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Longitude"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Latitude"
    End With
    With ActiveChart.Axes(xlCategory)
        .HasMajorGridlines = False
        .HasMinorGridlines = False
    End With
    With ActiveChart.Axes(xlValue)
        .HasMajorGridlines = False
        .HasMinorGridlines = False
    End With
    ActiveChart.PlotArea.Select
    Selection.ClearFormats
    ActiveChart.Axes(xlCategory).Select
    With ActiveChart.Axes(xlCategory)
        .MinimumScaleIsAuto = True
        .MaximumScaleIsAuto = True
        .MinorUnitIsAuto = True
        .MajorUnitIsAuto = True
        .Crosses = xlMaximum
        .ReversePlotOrder = False
        .ReversePlotOrder = True
        .ScaleType = xlLinear
        .DisplayUnit = xlNone
    End With

'###################### SELECT  Region_mapa TO FORMAT GRAPH

Region_mapa = Worksheets("Input").Rows(3).Columns(2)

Select Case Region_mapa

Case "Chile"
'Aca empieza el formateo de graficos para Loco

 ActiveChart.PlotArea.Select
    ActiveChart.Axes(xlValue).Select
    Selection.TickLabels.AutoScaleFont = False
    With Selection.TickLabels.Font
        .Name = "Arial"
        .FontStyle = "Regular"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .Background = xlAutomatic
    End With
    ActiveChart.Axes(xlCategory).Select
    Selection.TickLabels.AutoScaleFont = False
    With Selection.TickLabels.Font
        .Name = "Arial"
        .FontStyle = "Regular"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .Background = xlAutomatic
    End With
    ActiveChart.PlotArea.Select
    Selection.Width = 156
    Selection.Left = 472
    Selection.Width = 102
    Selection.Left = 526
    Selection.Width = 91
    Selection.Left = 537
    ActiveChart.Axes(xlValue).Select
    With ActiveChart.Axes(xlValue)
        .MinimumScaleIsAuto = True
        .MaximumScaleIsAuto = True
        .MinorUnitIsAuto = True
        .MajorUnitIsAuto = True
        .Crosses = xlAutomatic
        '.ReversePlotOrder = False
        .ReversePlotOrder = True
        .ScaleType = xlLinear
        .DisplayUnit = xlNone
    End With
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    Selection.AutoScaleFont = True
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "Bold"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .Background = xlAutomatic
    End With
    ActiveChart.Axes(xlValue).AxisTitle.Select
    Selection.AutoScaleFont = True
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "Bold"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .Background = xlAutomatic
    End With
    ActiveChart.Axes(xlValue).Select
    With ActiveChart.Axes(xlValue)
        
        '@@@@@@@@@@@@@@@ Start  for California Sea Urchin
                        
'        .MinimumScale = 32.6
'        .MaximumScale = 32.9
'        .MinorUnitIsAuto = True
'        .MajorUnitIsAuto = True
'        .Crosses = xlAutomatic
'        .ReversePlotOrder = True
'        .ScaleType = xlLinear
'        .DisplayUnit = xlNone
        
        '@@@@@@@@@@@@@@@ End  for California Sea Urchin
        
        
        .MinimumScale = 29
        .MaximumScale = 32.5

        .MinorUnitIsAuto = True
        .MajorUnitIsAuto = True
        .Crosses = xlAutomatic
        .ReversePlotOrder = True
        .ScaleType = xlLinear
        .DisplayUnit = xlNone
        
    End With

'aca termina el format nuevo para loco

'################################################################################################

Case "Puget_Sound"

'@@@@@@@@@@@@@@@@@@@  START      fomarting Geoducks from Puget Sound

    ActiveChart.PlotArea.Select
    Selection.Left = 150
    Selection.Width = 479
    ActiveChart.Axes(xlValue).Select
    With ActiveChart.Axes(xlValue)
        .MinimumScaleIsAuto = True
        .MaximumScaleIsAuto = True
        .MinorUnitIsAuto = True
        .MajorUnitIsAuto = True
        .Crosses = xlAutomatic
        .ReversePlotOrder = False
        .ScaleType = xlLinear
        .DisplayUnit = xlNone
    End With
    
    ActiveChart.Axes(xlCategory).Select
    With ActiveChart.Axes(xlCategory)
        .MinimumScaleIsAuto = True
        .MaximumScaleIsAuto = True
        .MinorUnitIsAuto = True
        .MajorUnitIsAuto = True
        .Crosses = xlMaximum
        .ReversePlotOrder = True
        .ScaleType = xlLinear
        .DisplayUnit = xlNone
    End With
    
    ActiveChart.Axes(xlCategory).Select
    With ActiveChart.Axes(xlCategory)
        .MinimumScaleIsAuto = True
        .MaximumScaleIsAuto = True
        .MinorUnitIsAuto = True
        .MajorUnitIsAuto = True
        .Crosses = xlMaximum
        .ReversePlotOrder = True
        .ScaleType = xlLinear
        .DisplayUnit = xlNone
    End With
        
    ActiveChart.SeriesCollection(1).Select
    With ActiveChart.ChartGroups(1)
        .VaryByCategories = False
        .ShowNegativeBubbles = False
        .SizeRepresents = xlSizeIsArea
        .BubbleScale = 10
    End With
'@@@@@@@@@@@@@@@@@@@  END       fomarting Geoducks from Puget Sound

End Select












2

Application.ScreenUpdating = True
End Sub

Sub clean(dd)
Attribute clean.VB_ProcData.VB_Invoke_Func = " \n14"
  On Error Resume Next
  Dim file_name As String
  Application.DisplayAlerts = False
     On Error Resume Next
     Sheets(dd).Select
     On Error Resume Next
     Sheets(dd).Delete
     
End Sub

Sub referenciagraph()
Attribute referenciagraph.VB_ProcData.VB_Invoke_Func = " \n14"
Dim j As Integer, Area As Integer, k As Double
Dim VarEspacial2 As Integer, ScaleFlag As Integer
    
VarEspacial2 = Worksheets("Input").Rows(8).Columns(3)
ScaleFlag = Worksheets("Input").Rows(8).Columns(4)
Nyears = Worksheets("Input").Rows(33).Columns(2) - Worksheets("Input").Rows(32).Columns(2) + 1
Nareas = Worksheets("Input").Rows(31).Columns(2)

    Application.ScreenUpdating = False
For j = 1 To Nyears

    Worksheets("Graphs").Rows(3).Columns(99) = Worksheets("Output").Rows(1 + j).Columns(99)
    
    For Area = 1 To Nareas
        Select Case ScaleFlag
            Case 1
            Worksheets("Graphs").Rows(3).Columns(99 + Area) = _
            (Worksheets("Output").Rows(1 + j + ((VarEspacial2 - 1) * Nyears)).Columns(Area + 99))

            Case 2
            Worksheets("Graphs").Rows(3).Columns(99 + Area) = _
            Sqr(Worksheets("Output").Rows(1 + j + ((VarEspacial2 - 1) * Nyears)).Columns(Area + 99))

            Case 3
            Worksheets("Graphs").Rows(3).Columns(99 + Area) = _
            Log(Worksheets("Output").Rows(1 + j + ((VarEspacial2 - 1) * Nyears)).Columns(Area + 99))
        End Select
    Next

    For k = 1 To 1000
        Worksheets("Graphs").Rows(1).Columns(1) = k ^ 12
    Next
        Application.ScreenUpdating = True

Next

End Sub

