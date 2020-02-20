Attribute VB_Name = "Módulo3"
Sub EliminarHoja()
Attribute EliminarHoja.VB_ProcData.VB_Invoke_Func = " \n14"
    On Error Resume Next
    Sheets("Detalle Pozos").Delete
    
End Sub
Sub AñadirPAD()
    
    UserForm2.Show
    
End Sub
Sub AñadirPozo()

    UserForm3.Show

End Sub
Sub AñadirPAD2()
    
    UserForm4.Show
    
End Sub
Sub AñadirEquipos()
    
    UserForm5.Show
    
End Sub
