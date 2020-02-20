VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "UserForm4"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2985
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Ingresar_Click()
    
    Application.ScreenUpdating = False
    If Trim(TextBox1.Text) <> "" Then
        For cuenta = 1 To 12000
            If Sheets("Base.Prod").Cells(cuenta, 2) = "FINAL" Then dato = cuenta
        Next cuenta
        
        For cuenta1 = dato + 4 To 12000
            If Sheets("Base.Prod").Cells(cuenta1, 2) = "FINAL CONSUMO" Then datofinal = cuenta1
        Next cuenta1
        
        Rows(datofinal & ":" & datofinal).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Rows(datofinal & ":" & datofinal).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        'Selecciona nueva fila creada
        Sheets("Base.Prod").Cells(datofinal, 2) = TextBox1.Value
        Sheets("Base.Prod").Cells(datofinal + 1, 2) = "Horas de Funcionamiento"
        
        With Sheets("Base.Prod").Cells(datofinal, 2).Font
            .Name = "Calibri"
            .Size = 11
        End With
        
        With Sheets("Base.Prod").Cells(datofinal + 1, 2).Font
            .Name = "Calibri"
            .Size = 9
        End With
        
        With Range(Cells(datofinal, 2), Cells(datofinal, 381)).Borders
            .LineStyle = xlContinuous
        End With
        With Range(Cells(datofinal + 1, 2), Cells(datofinal, 381)).Borders
            .LineStyle = xlContinuous
        End With
        MsgBox TextBox1.Value + " se ingresó exitosamente."
        
        Unload UserForm4
    Else
        Label2.Visible = True
    End If
End Sub
