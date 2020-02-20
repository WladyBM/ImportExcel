Attribute VB_Name = "Módulo1"
Sub Importar()

Dim Archivo As String
Set Selector = Application.FileDialog(msoFileDialogOpen)

    With Selector
        .Title = "Por favor seleccione el archivo"
        If .Show <> -1 Then
            Exit Sub
        End If
        fileselected = .SelectedItems(1)
        Archivo = fileselected
    End With
    
    Call ListarHojas(Archivo)
    
    End Sub
    
    
Sub ListarHojas(Archivo As String)
    Dim Hoja As Worksheet

    UserForm1.lblRutaArchivo = Archivo
    UserForm1.cmbListadoHojas.Clear
    UserForm1.cmbListadoHojas.AddItem "*todas"
    
    Workbooks.Open (Dir(Archivo))
    
    For Each Hoja In Workbooks(Dir(Archivo)).Worksheets
        UserForm1.cmbListadoHojas.AddItem Hoja.Name
    Next Hoja

    Workbooks(Dir(Archivo)).Close
    UserForm1.Show

End Sub

Sub ImportarHojaExcel(Archivo As String, hojaimportar As String)
    Dim carpeta As String, Nombre_archivo As String, Nombre_archivo_actual As String, Hoja As Worksheet, total As Integer
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Nombre_archivo_actual = ThisWorkbook.Name
    Nombre_archivo = Dir(Archivo)
    
    If Nombre_archivo = "" Then
        Exit Sub
    End If
    
    Workbooks.Open (Nombre_archivo)
    If hojaimportar = "*todas" Then
        For Each Hoja In Workbooks(Nombre_archivo).Worksheets
            total = Workbooks(Nombre_archivo_actual).Worksheets.Count
            Workbooks(Nombre_archivo).Worksheets(Hoja.Name).Copy _
            after:=Workbooks(Nombre_archivo_actual).Worksheets(total)
        Next Hoja
    Else
        total = Workbooks(Nombre_archivo_actual).Worksheets.Count
        Workbooks(Nombre_archivo).Worksheets(hojaimportar).Copy _
        after:=Workbooks(Nombre_archivo_actual).Worksheets(total)
    End If
    
    Workbooks(Nombre_archivo).Close
    
    ThisWorkbook.Sheets("Base.Prod").Activate
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Hoja importada con éxito"
    
    Unload UserForm1
End Sub

