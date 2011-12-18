Attribute VB_Name = "M0_Main"
Option Explicit
Dim year As Integer, t As Integer, monte As Integer

Sub Main()
Attribute Main.VB_ProcData.VB_Invoke_Func = " \n14"

Worksheets("Time").Rows(1).Columns(2) = Now

Application.ScreenUpdating = False

Call Read_Input.Read_Input
Call Preliminary_Calcs.Initialize_variables
Call Input_Output.Output_Initialize

Call Preliminary_Calcs.Set_Carrying_Capacity
Call Preliminary_Calcs.Set_Virgin_Conditions     'NB Virgin conditions may be below K-  if Lambda < 1 some areas are limited by larval supply
Call Preliminary_Calcs.Set_InitialConditions
     
Select Case RunFlags.Run_type

Case 1  'conditioning run
    Call M2_Conditioning.Conditioning
    Call M2_Conditioning.FitData
    Call M2_Conditioning.CalcLikelihood

Case Else   'simulation run
  'Both in the conditioning and preliminary calcs the dynamic is annual
  'even in cases of intrayear dynamics, therefore rescale parameters is done here
  'before simmulation
  Worksheets("Time").Rows(3).Columns(2) = Nreplicates
  Application.DisplayStatusBar = True
  Call Preliminary_Calcs.Rescale_parameters
  
  For monte = 1 To Nreplicates
        
   Application.StatusBar = "Runnning simmulation " & monte & " out of " & Nreplicates
        
        iz = 1 + (monte - 1) * (N_Zvector / Nreplicates)
    Call M2_Random_Stuff.VariableInitialConditions
    Call M2_Random_Stuff.RecruitmentDevs

    For year = StYear To EndYear
        Call M4_Calc_Recruits.Random_Recruits(year)
            If Nreplicates = 1 Then Call M2_AnnualUpdate.pLgen(year)
        
        Call M6_Prod_Alloc_Larvae.Prod_Alloc_Larvae(year, Bmature)
        
        Call Management_Procedure.Strategies(year)
         
        'Aca entra para la dinamica intraanual, si es solo anual entonces el loop es phony
        For t = 1 To Nt
             
            If OpenMonth(year, t) = True Then
            
                Call M7_Fishing.Fishing(year, t)
            
            Else
            
                For Area = 1 To Nareas
                    HRTmp(Area) = 0
                Next Area
            
            End If
            
            Call M5_Popdyn.PopDyn(year)
            
        Next t
        
        Call M2_AnnualUpdate.Annual_update(year)
  
    Next year
  
    Call Input_Output.Print_Output(monte)
    
    Application.ScreenUpdating = True
    Application.ScreenUpdating = False
    If RunFlags.Output_NAge_NSize = True Then Call Input_Output.Print_Output_NAge_NSize(monte)
  
  Next monte
    
    Call Input_Output.Print_Input
    Call Graph.Graphs
    If RunFlags.Output_Size_W = True Then Call Input_Output.Print_Size_W
    
Worksheets("Time").Rows(2).Columns(2) = Now

Application.ScreenUpdating = True
Application.StatusBar = False
End Select

End 'To reset all module-level variables of all modules

End Sub
