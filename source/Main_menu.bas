Attribute VB_Name = "Main_menu"
Option Explicit
Sub Makemenu()
Attribute Makemenu.VB_ProcData.VB_Invoke_Func = " \n14"
 Dim NewMenuBar As CommandBar
 Dim MyPath As String
 Dim dummy As String
 '##########################
 '#         SET VERSION NAME      #
 '##########################
 
 
 'Delete menu bar if it exists
    Call DeleteMenuBar
 
 'Add a menu bar
 Set NewMenuBar = CommandBars.Add(MenuBar:=True)
    With NewMenuBar
        .Name = "MyMenuBar"
        .Visible = True
    End With

'Copy the File menu from Worksheet Menu Bar
CommandBars("Worksheet Menu Bar").Controls(1).Copy bar:=CommandBars("MyMenuBar")

 'Add new menu item (MetaPesca)
 
 Dim MyMenu As CommandBarPopup
 With Application
  
    Set MyMenu = .CommandBars("MyMenuBar").Controls.Add(msoControlPopup, temporary:=True)
    With MyMenu
      .Caption = "&MetaPesca"
    End With
    
    Language = Worksheets("TBSheet").Rows(1).Columns(2)
    
    
'###########################################
'#                           MAIN MENUS                                         #
'###########################################
 
    Dim OpenFile As CommandBarButton
    Dim NewFile As CommandBarButton
    Dim Run As CommandBarButton
    Dim Help As CommandBarButton
    Dim EditCode As CommandBarButton
    Dim Restore As CommandBarButton
    Dim ShortMenu As CommandBarButton
    Dim Helpabout As CommandBarButton
    Dim Conectividad As CommandBarButton
    Dim Manejo As CommandBarButton
'    Dim Pesca As CommandBarButton
    Dim OutputOptions As CommandBarButton
    Dim PopDyn As CommandBarButton
    Dim Exportar_dat As CommandBarButton
    Dim Condiciones_0 As CommandBarButton
    Dim Zoom As CommandBarButton
    
    Set OpenFile = MyMenu.Controls.Add(msoControlButton)
        With OpenFile
            If Language = "English" Then
                .Caption = "&Open File"
            Else
                .Caption = "&Abrir Archivo"
            End If
            .OnAction = "GotoForm1"
        End With
      
      Set Condiciones_0 = MyMenu.Controls.Add(msoControlButton)
        With Condiciones_0
            If Language = "English" Then
                .Caption = "&Initial Conditions"
            Else
                .Caption = "&Condiciones Iniciales"
            End If
            .OnAction = "GotoCondicionesIniciales"
        End With
      
    Set NewFile = MyMenu.Controls.Add(msoControlButton)
        With NewFile
            If Language = "English" Then
                .Caption = "&Create New File"
            Else
                .Caption = "&Crear Archivo Nuevo"
            End If
         .OnAction = "GotoForm2"
        End With
    
    Set PopDyn = MyMenu.Controls.Add(msoControlButton)
        With PopDyn
            If Language = "English" Then
                .Caption = "&Population Dynamics"
            Else
                .Caption = "&Dinamica Poblacional"
            End If
         
         .OnAction = "ShowPopDyn"
        End With
       
     Set Conectividad = MyMenu.Controls.Add(msoControlButton)
        With Conectividad
         
            If Language = "English" Then
                .Caption = "&Larval Connectivity"
            Else
                .Caption = "&Conectividad Larvaria"
            End If
         .OnAction = "Showconectividad"
        End With
   
 ' Set Pesca = MyMenu.Controls.Add(msoControlButton)
  '      With Pesca
   '      If Language = "English" Then
    '            .Caption = "&Fisheries Dynamics"
     '       Else
      '          .Caption = "&Dinamica Pesquera"
       '     End If
        ' .OnAction = "Showpesca"
        'End With
      
   Set Manejo = MyMenu.Controls.Add(msoControlButton)
        With Manejo
         
         If Language = "English" Then
                .Caption = "&Management"
            Else
                .Caption = "&Manejo"
            End If
         .OnAction = "Showmanagement"
        End With
   
   Set OutputOptions = MyMenu.Controls.Add(msoControlButton)
        With OutputOptions
         
         If Language = "English" Then
                .Caption = "&Output Options"
            Else
                .Caption = "&Opciones de Salida"
            End If
         .OnAction = "ShowOutputOptions"
        End With
   
   Set Run = MyMenu.Controls.Add(msoControlButton)
        With Run
         
         If Language = "English" Then
                .Caption = "&RUN MODEL"
            Else
                .Caption = "&Iniciar simulacion"
            End If
         .OnAction = "M0_Main.Main"
         
        End With
      
    Set EditCode = MyMenu.Controls.Add(msoControlButton)
        With EditCode
         
         If Language = "English" Then
                .Caption = "&Edit Code Alt+F11"
            Else
                .Caption = "&Editar Codigo Alt+F11"
            End If
         .OnAction = "CodeMessage"
        End With
    
    Set Exportar_dat = MyMenu.Controls.Add(msoControlButton)
        With Exportar_dat
         
         If Language = "English" Then
                .Caption = "&Export Input File"
            Else
                .Caption = "&Exportar Archivo"
            End If
         .OnAction = "Goto_Export_dat"
        End With
            
    Set Restore = MyMenu.Controls.Add(msoControlButton)
    With Restore
        
        If Language = "English" Then
                .Caption = "&Restore Excel Menu"
            Else
                .Caption = "&Restaurar Menu de Excel"
            End If
        .OnAction = "RestoreExcelMenuMetapesca"
    End With
    
    Set Restore = MyMenu.Controls.Add(msoControlButton)
    With Restore
        
        If Language = "English" Then
                .Caption = "&Restore MetaPesca Menu"
            Else
                .Caption = "&Restaurar Menu de Metapesca"
            End If
        .OnAction = "Makemenu"
    End With
    
    Set Helpabout = MyMenu.Controls.Add(msoControlButton)
    With Helpabout
        
        If Language = "English" Then
                .Caption = "&About Metapesca"
            Else
                .Caption = "&Sobre Metapesca"
            End If
        .OnAction = "Gotoform3"
    End With
    
    Set Help = MyMenu.Controls.Add(msoControlButton)
        With Help
         
         If Language = "English" Then
                .Caption = "&Help"
            Else
                .Caption = "&Ayuda"
            End If
         .HyperlinkType = msoCommandBarButtonHyperlinkNone
        End With
        dummy = ThisWorkbook.Path
        MyPath = dummy & "\InputMetapesca26p.doc"
 
        If Help.HyperlinkType <> _
            msoCommandBarButtonHyperlinkOpen Then
            Help.HyperlinkType = _
            msoCommandBarButtonHyperlinkOpen
            Help.TooltipText = MyPath
        End If
     
  End With
    
    Set Zoom = MyMenu.Controls.Add(msoControlButton)
        With Zoom
         
         If Language = "English" Then
                .Caption = "&Zoom"
            Else
                .Caption = "&Zoom"
            End If
         .OnAction = "Goto_Zoom"
        End With


End Sub
Sub Removemenu()
Attribute Removemenu.VB_ProcData.VB_Invoke_Func = " \n14"
On Error Resume Next
  CommandBars("worksheet menu bar").Controls("MetaPesca").Delete
  'MyMenu.Remove
End Sub
Sub DeleteMenuBar()
Attribute DeleteMenuBar.VB_ProcData.VB_Invoke_Func = " \n14"
    On Error Resume Next
    CommandBars("MyMenuBar").Delete
    On Error GoTo 0
End Sub
Sub GotoForm1()
Attribute GotoForm1.VB_ProcData.VB_Invoke_Func = " \n14"
  Main_Form.Show
End Sub
Sub CodeMessage()
Attribute CodeMessage.VB_ProcData.VB_Invoke_Func = " \n14"
    MsgBox ("Press Alt + F11 to Edit Code")
End Sub
Sub GotoForm2()
Attribute GotoForm2.VB_ProcData.VB_Invoke_Func = " \n14"
  NewFileForm.Show
End Sub
Sub GotoForm3()
Attribute GotoForm3.VB_ProcData.VB_Invoke_Func = " \n14"
  About.Show
End Sub
Sub RestoreExcelMenuMetapesca()
Attribute RestoreExcelMenuMetapesca.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim i As Integer
    For i = 2 To 10
    CommandBars("Worksheet Menu Bar").Controls(i).Copy bar:=CommandBars("MyMenuBar")
    Next i
End Sub
Sub Showconectividad()
Attribute Showconectividad.VB_ProcData.VB_Invoke_Func = " \n14"
  Conectividad.Show
End Sub
Sub Showmanagement()
Attribute Showmanagement.VB_ProcData.VB_Invoke_Func = " \n14"
  MsgBox ("Please modify management options from 'Input' Sheet")
 ' InputManagement.Show
End Sub
'Sub Showpesca()
'    Pesca.Show
'End Sub
Sub ShowOutputOptions()
Attribute ShowOutputOptions.VB_ProcData.VB_Invoke_Func = " \n14"
  Output.Show
End Sub
Sub ShowPopDyn()
Attribute ShowPopDyn.VB_ProcData.VB_Invoke_Func = " \n14"
    PopDyn.Show
End Sub
Sub Goto_Export_dat()
Attribute Goto_Export_dat.VB_ProcData.VB_Invoke_Func = " \n14"
    Main_Form.Export_dat
End Sub
Sub GotoCondicionesIniciales()
Attribute GotoCondicionesIniciales.VB_ProcData.VB_Invoke_Func = " \n14"
    Initial_Conditions.Show
End Sub
Sub Goto_Zoom()
Attribute Goto_Zoom.VB_ProcData.VB_Invoke_Func = " \n14"
    Zoom.Show
End Sub
