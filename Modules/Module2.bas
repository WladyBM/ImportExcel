Attribute VB_Name = "Módulo2"
Sub Transferir()
    Dim datofinal1 As Long
    Dim datofinal2 As Long
    Dim Celda1 As Object
    Dim Ubicacion As String
    Dim DetallePozo As Range
    
    Application.ScreenUpdating = False
    
    'Cuenta el número de filas hasta encontrar aquella que posee el dato INICIO PRODUCCION y lo transfiere a la variable dato
    For cuenta1 = 1 To 12000
        If Sheets("Base.Prod").Cells(cuenta1, 2) = "INICIO PRODUCCION" Then dato = cuenta1
    Next cuenta1
    
    'Cuenta el numero de filas en Detalle Pozos hasta el que dice TOTAL BLOQUE, así darle el valor a la variable datofinal2
    For cuenta2 = 12 To 1000
        On Error Resume Next
        If Sheets("Detalle Pozos").Cells(cuenta2, 5) = "TOTAL BLOQUE :" Then
            datofinal2 = cuenta2
            Exit For
        ElseIf Sheets("Detalle Pozos").Cells(cuenta2, 5) = "DETALLE POZOS CON EQUIPOS SLA (m3/d)" Then
            datofinal2 = cuenta2 - 2
        End If
        If Err.Number = 9 Then
            MsgBox "Añada la hoja Detalle Pozos para transferir datos."
            Exit Sub
        End If
    Next cuenta2
    
    'Inserta un "_" a partir de la última columna que tenga ingresado un valor
    Sheets("Base.prod").Cells(dato, Columns.Count).End(xlToLeft).Offset(0, 1) = "_"
    
    'Cuenta el número de columnas hasta encontrar aquella que posee un dato, en este aquella que tiene el "_"
    ultima_columna = Sheets("Base.prod").Cells(dato, Columns.Count).End(xlToLeft).Column
    
    'Nuevamente cuenta el numero de filas hasta encontrar aquella que tiene el dato FINAL y es usado para establecer el final donde están los pozos ingresados
    For cuenta1 = dato + 2 To 12000
        If Sheets("Base.Prod").Cells(cuenta1, 2) = "FINAL" Then datofinal1 = cuenta1
    Next cuenta1
    
    'Establece una direccion llamada DetallePozo para poder leer cada celda que los componga
    Set DetallePozo = Sheets("Detalle Pozos").Range("E12:E" & datofinal2)
    
    'Le cada celda por medio de un bucle
    For Each Celda1 In DetallePozo
        Ubicacion = Celda1.Offset(0, 10).Value
        'Si encuentra un dato que coincida en la Hoja Detalle Produccion y Detalle Pozo, empieza a añadir los datos
        For Fila = dato + 2 To datofinal1
            On Error Resume Next
            If Sheets("Base.Prod").Cells(Fila, 3) Like Celda1 Then
                Sheets("Base.Prod").Cells(Fila, ultima_columna).Select
                ActiveCell.Offset(0, 0) = Ubicacion * 1000
            ElseIf Sheets("Base.Prod").Cells(Fila, 2) Like Celda1 Then
                Sheets("Base.Prod").Cells(Fila, ultima_columna) = Ubicacion * 1000
            End If
        Next Fila
    Next Celda1
    Transferir2
    
End Sub
Sub Transferir2()
    Dim dato As Long
    Dim datofinal1 As Long
    Dim datofinal2 As Long
    Dim Celda1 As Object
    Dim Hour As Long
    Dim DetallePozo As Range
    
    Application.ScreenUpdating = False
    
    'Validar donde iniciar el registro de datos a partir de la celda que dice FINAL, se usa como punto de inicio
    For cuenta1 = 1 To 12000
        If Sheets("Base.Prod").Cells(cuenta1, 2) = "FINAL" Then dato = cuenta1
    Next cuenta1
    
    'Cuenta el numero de filas a partir de FINAL hasta FINAL CONSUMO
    For cuenta1 = dato + 6 To 12000
        If Sheets("Base.Prod").Cells(cuenta1, 2) = "FINAL CONSUMO" Then datofinal1 = cuenta1
    Next cuenta1
    For cuenta2 = 12 To 1000
        On Error Resume Next
        If Sheets("Detalle Pozos").Cells(cuenta2, 5) = "TOTAL BLOQUE :" Then
            datofinal2 = cuenta2
        ElseIf Sheets("Detalle Pozos").Cells(cuenta2, 5) = "DETALLE POZOS CON EQUIPOS SLA (m3/d)" Then
            datofinal2 = cuenta2 - 2
        End If
        If Err.Number = 9 Then
            MsgBox "Añada la hoja Detalle Pozos para transferir datos."
            Exit Sub
        End If
    Next cuenta2
    
    'Establece la columna a ocupar insertando un "_" a partir de la ultima columna sin datos ingresados
    Sheets("Base.prod").Cells(dato + 4, Columns.Count).End(xlToLeft).Offset(0, 1) = "_"
    
    'Cuenta la ultima columna con datos, la que posee la "_"
    ultima_columna = Sheets("Base.prod").Cells(dato + 4, Columns.Count).End(xlToLeft).Column
    
    'Establece la direccion DetallePozo, para posteriormente usarla para contar cada celda que lo compone
    Set DetallePozo = Sheets("Detalle Pozos").Range("E12:E" & datofinal2)
    
    'Lee cada celda en un bucle
    For Each Celda1 In DetallePozo
        Hora = Celda1.Offset(0, 1).Value
        'Si coinciden los PAD o pozos en Detalle Pozos y Detalle Produccion, añade las horas
        For Fila = dato + 6 To datofinal1
            If IsEmpty(Sheets("Base.Prod").Cells(Fila, 2)) Then
            Else
                If Celda1 Like Sheets("Base.Prod").Cells(Fila, 2) & "*" Then
                    If IsEmpty(Sheets("Base.Prod").Cells(Fila + 1, ultima_columna)) Then
                        Sheets("Base.Prod").Cells(Fila + 1, ultima_columna) = Hora
                    Else
                        If Sheets("Base.Prod").Cells(Fila + 1, ultima_columna).Value > Hora Then
                        ElseIf Sheets("Base.Prod").Cells(Fila + 1, ultima_columna).Value = Hora Then
                        Else
                            Sheets("Base.Prod").Cells(Fila + 1, ultima_columna) = Hora
                        End If
                    End If
                ElseIf Sheets("Base.Prod").Cells(Fila, 2) Like Celda1 Then
                    If IsEmpty(Sheets("Base.Prod").Cells(Fila + 1, ultima_columna)) Then
                        Sheets("Base.Prod").Cells(Fila + 1, ultima_columna) = Hora
                    Else
                        If Sheets("Base.Prod").Cells(Fila + 1, ultima_columna).Value > Hora Then
                        ElseIf Sheets("Base.Prod").Cells(Fila + 1, ultima_columna).Value = Hora Then
                        Else
                            Sheets("Base.Prod").Cells(Fila + 1, ultima_columna) = Hora
                        End If
                    End If
                End If
            End If
        Next Fila
    Next Celda1
    
    '-------------------- Seccion de filtrado para Calentadores ------------------------------
    'Verifica cada celda que posee inicia con la palabra Calentador, de modo para ingresar un valor vacío
    For Fila = dato + 6 To datofinal1
        If Sheets("Base.Prod").Cells(Fila, 3) Like "Calentador*" Then
            Sheets("Base.Prod").Cells(Fila, ultima_columna) = ""
    
    '-------------------- Seccion de filtrado para Generadores ------------------------------
    'Verifica cada celda que posea la palabra Generador, para ingresar una formula de manera predeterminada
        ElseIf Sheets("Base.Prod").Cells(Fila, 3) Like "Generador*" Then
            For Referencial = Fila To dato + 6 Step -1
                If Sheets("Base.Prod").Cells(Referencial, 2) = "Horas de Funcionamiento" Then
                   Rango1 = Sheets("Base.Prod").Cells(Referencial, ultima_columna).Address(False, False)
                   Exit For
                End If
            Next Referencial
            
            Rango2 = Sheets("Base.Prod").Cells(Fila, ultima_columna).Address(False, False)
        'Verifica si las celdas coincidentes poseen datos ingresados o estan vacios
            If Not IsEmpty(Range(Rango2)) Then
            Else
                Range(Rango2).Formula = "=(197/24)*" & Rango1
            End If
    
    '-------------------- Seccion de filtrado para equipos URG Q.B.JOHNSON* ------------------------------
    'Verifica cada celda que posea la palabra URG, para ingresar otra formula predeterminada
        ElseIf Sheets("Base.Prod").Cells(Fila, 3) Like "URG Q.B.JOHNSON*" Then
            For Referencial = Fila To dato + 6 Step -1
                If Sheets("Base.Prod").Cells(Referencial, 2) = "Horas de Funcionamiento" Then
                   Rango1 = Sheets("Base.Prod").Cells(Referencial, ultima_columna).Address(False, False)
                   Exit For
                End If
            Next Referencial
            
            Rango2 = Sheets("Base.Prod").Cells(Fila, ultima_columna).Address(False, False)
        'Verifica si las celdas coincidentes poseen datos ingresados o estan vacios
            If Not IsEmpty(Range(Rango2)) Then
            Else
                Range(Rango2).Formula = "=(178.9/24)*" & Rango1
            End If
            
        '-------------------- Seccion de filtrado para equipos URG 500,000 BTU ------------------------------
    'Verifica cada celda que posea la palabra URG, para ingresar otra formula predeterminada
        ElseIf Sheets("Base.Prod").Cells(Fila, 3) Like "URG 500,000 BTU" Then
            For Referencial = Fila To dato + 6 Step -1
                If Sheets("Base.Prod").Cells(Referencial, 2) = "Horas de Funcionamiento" Then
                   Rango1 = Sheets("Base.Prod").Cells(Referencial, ultima_columna).Address(False, False)
                   Exit For
                End If
            Next Referencial
            
            Rango2 = Sheets("Base.Prod").Cells(Fila, ultima_columna).Address(False, False)
        'Verifica si las celdas coincidentes poseen datos ingresados o estan vacios
            If Not IsEmpty(Range(Rango2)) Then
            Else
                Range(Rango2).Formula = "=(447.3/24)*" & Rango1
            End If
         
        '-------------------- Seccion de ejemplo para copiar y pega ------------------------------
        
    'Verifica cada celda que posea la palabra Motogenerador, para ingresar otra formula predeterminada
        ElseIf Sheets("Base.Prod").Cells(Fila, 3) Like "Motogenerador*" Then
            For Referencial = Fila To dato + 6 Step -1
                If Sheets("Base.Prod").Cells(Referencial, 2) = "Horas de Funcionamiento" Then
                   Rango1 = Sheets("Base.Prod").Cells(Referencial, ultima_columna).Address(False, False)
                   Exit For
                End If
            Next Referencial
        
            Rango2 = Sheets("Base.Prod").Cells(Fila, ultima_columna).Address(False, False)
        'Verifica si las celdas coincidentes poseen datos ingresados o estan vacios
            If Not IsEmpty(Range(Rango2)) Then
            Else
                Range(Rango2).Formula = "=(197/24)*" & Rango1
            End If
            
        '-------------------- Seccion de ejemplo para copiar y pega ------------------------------
        
    'Verifica cada celda que posea la palabra URG, para ingresar otra formula predeterminada
        'ElseIf Sheets("Base.Prod").Cells(Fila, 3) Like "NOMBRE EQUIPO" Then
            'For Referencial = Fila To dato + 6 Step -1
                'If Sheets("Base.Prod").Cells(Referencial, 2) = "Horas de Funcionamiento" Then
                   'Rango1 = Sheets("Base.Prod").Cells(Referencial, ultima_columna).Address(False, False)
                   'Exit For
                'End If
            'Next Referencial
        
            'Rango2 = Sheets("Base.Prod").Cells(Fila, ultima_columna).Address(False, False)
        'Verifica si las celdas coincidentes poseen datos ingresados o estan vacios
            'If Not IsEmpty(Range(Rango2)) Then
            'Else
                'Range(Rango2).Formula = "=FORMULA" & Rango1
            'End If
              
        End If
    Next Fila
    'Eliminar hoja
    Call Módulo3.EliminarHoja
    
End Sub
