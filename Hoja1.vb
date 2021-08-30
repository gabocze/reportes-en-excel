Dim MyTarget
Public msg As String
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Llenar_TextBox1 Target
    Llenar_TextBox2 Target
    MyTarget = Target
    msg = "mensaje global"
End Sub

Sub Llenar_TextBox1(ByVal Target As Range)
   UserForm1.TextBox1.Text = Worksheets(1).Name
   UserForm1.TextBox1.Text = UserForm1.TextBox1.Text + CStr(Columna1_Fila1(Target))
End Sub

Sub Llenar_TextBox2(ByVal Target As Range)
    iCnt = 0
    For Each c In Target
        iCnt = iCnt + 1
    Next c
    UserForm1.TextBox2.Text = "TOTAL: " + CStr(iCnt) + "; " + "Filas: " + CStr(Target.Rows.Count) + ", " + "Columnas: " + CStr(Target.Columns.Count) + "; " + "Fila y Columna inicial: " + CStr(Target.Row) + ", " + CStr(Target.Column) + "."
End Sub

Public Sub Copiar_Rango()
    'Ubica la fila inicial del destino como la Fila final del rango original + 4
    fila_inicial_origen = MyTarget.Row
    fila_inicial_destino = fila_inicial_origen + MyTarget.Rows.Count + 3
    rango = Range("A6").Select
    Range("A6", "F9").Select
    MyTarget.Copy
    'Paste 'anda bien
End Sub

Function Columna1_Fila1(ByVal Target As Range)
On Error GoTo Manejo_Error
    Columna1_Fila1 = Target.Value2(1, 1) 'Las fechas se muestran como números
    Exit Function
Manejo_Error:
    Columna1_Fila1 = Columna1_Fila1_Como_NoRango(Target)
End Function

Function Columna1_Fila1_Como_NoRango(ByVal Target As Range)
    Columna1_Fila1_Como_Rango = Target.Value 'Las fechas se muestran como fechas
End Function
