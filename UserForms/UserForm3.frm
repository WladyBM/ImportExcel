VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "UserForm3"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2910
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Ingresar_Click()
    Application.ScreenUpdating = False
    If Trim(TextBox1.Text) <> "" Then
        Dim datofinal As Integer
        
        For cuenta = 1 To 12000
            If Sheets("Base.Prod").Cells(cuenta, 2) = "INICIO PRODUCCION" Then dato = cuenta
        Next cuenta
        
        For cuenta1 = dato To 12000
            If Sheets("Base.Prod").Cells(cuenta1, 2) = "FINAL" Then datofinal = cuenta1
        Next cuenta1
        'Selecciona fila tomando referencia la que dice FINAL
        Rows(datofinal & ":" & datofinal).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        'Selecciona nueva fila creada
        Sheets("Base.Prod").Cells(datofinal, 3) = TextBox1.Value
        
        With Sheets("Base.Prod").Cells(datofinal, 3).Font
            .Name = "Calibri"
            .Size = 10
            .Color = RGB(128, 128, 128)
        End With
        
        MsgBox TextBox1.Value + " se ingres� exitosamente."
        
        Unload UserForm3
    Else
        Label2.Visible = True
    End If
End Sub
