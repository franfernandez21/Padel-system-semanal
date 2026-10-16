Attribute VB_Name = "Módulo1"
Sub GenerarFixtureAmericanoSeparado()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Leer los nombres de los jugadores de las celdas A2 a A9
    Dim jugadores(1 To 8) As String
    Dim i As Integer
    For i = 1 To 8
        If ws.Cells(i + 1, 1).Value = "" Then
            MsgBox "Faltan jugadores. Asegurate de tener 8 nombres en la columna A (A2:A9).", vbExclamation
            Exit Sub
        End If
        jugadores(i) = ws.Cells(i + 1, 1).Value
    Next i
    
    ' Matriz predefinida con las 7 fechas (los números del 1 al 8)
    Dim fixture As Variant
    fixture = Array( _
        Array(1, 2, 3, 4, 5, 6, 7, 8), _
        Array(1, 3, 5, 7, 2, 4, 6, 8), _
        Array(1, 4, 6, 7, 2, 3, 5, 8), _
        Array(1, 5, 2, 8, 3, 7, 4, 6), _
        Array(1, 6, 2, 5, 3, 8, 4, 7), _
        Array(1, 7, 4, 8, 2, 6, 3, 5), _
        Array(1, 8, 2, 7, 3, 6, 4, 5) _
    )
    
    ' Configurar los nuevos encabezados extendidos en las columnas C a M
    ws.Cells(1, 3).Value = "Fecha"
    ws.Cells(1, 4).Value = "Partido"
    ws.Cells(1, 5).Value = "N°"
    ws.Cells(1, 6).Value = "Jugador"
    ws.Cells(1, 7).Value = "N°"
    ws.Cells(1, 8).Value = "Jugador"
    ws.Cells(1, 9).Value = "VS"
    ws.Cells(1, 10).Value = "N°"
    ws.Cells(1, 11).Value = "Jugador"
    ws.Cells(1, 12).Value = "N°"
    ws.Cells(1, 13).Value = "Jugador"
    
    ' Formato para los encabezados
    With ws.Range("C1:M1")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(31, 78, 120) ' Color azul oscuro de fondo
        .Font.Color = RGB(255, 255, 255) ' Texto blanco
    End With
    
    Dim fila As Integer
    fila = 2
    Dim f As Integer
    Dim j1 As Integer, j2 As Integer, j3 As Integer, j4 As Integer
    
    ' Bucle para imprimir las 7 fechas con celdas separadas
    For f = 0 To 6
        ' ---------------- PARTIDO 1 ----------------
        j1 = fixture(f)(0): j2 = fixture(f)(1): j3 = fixture(f)(2): j4 = fixture(f)(3)
        ws.Cells(fila, 3).Value = "Fecha " & (f + 1)
        ws.Cells(fila, 4).Value = "Partido 1"
        ws.Cells(fila, 5).Value = j1
        ws.Cells(fila, 6).Value = jugadores(j1)
        ws.Cells(fila, 7).Value = j2
        ws.Cells(fila, 8).Value = jugadores(j2)
        ws.Cells(fila, 9).Value = "vs"
        ws.Cells(fila, 10).Value = j3
        ws.Cells(fila, 11).Value = jugadores(j3)
        ws.Cells(fila, 12).Value = j4
        ws.Cells(fila, 13).Value = jugadores(j4)
        fila = fila + 1
        
        ' ---------------- PARTIDO 2 ----------------
        j1 = fixture(f)(4): j2 = fixture(f)(5): j3 = fixture(f)(6): j4 = fixture(f)(7)
        ws.Cells(fila, 3).Value = "Fecha " & (f + 1)
        ws.Cells(fila, 4).Value = "Partido 2"
        ws.Cells(fila, 5).Value = j1
        ws.Cells(fila, 6).Value = jugadores(j1)
        ws.Cells(fila, 7).Value = j2
        ws.Cells(fila, 8).Value = jugadores(j2)
        ws.Cells(fila, 9).Value = "vs"
        ws.Cells(fila, 10).Value = j3
        ws.Cells(fila, 11).Value = jugadores(j3)
        ws.Cells(fila, 12).Value = j4
        ws.Cells(fila, 13).Value = jugadores(j4)
        fila = fila + 1
    Next f
    
    
    ' Centrar las columnas de los números y el "VS" para que quede más prolijo
    ws.Columns("E:E").HorizontalAlignment = xlCenter
    ws.Columns("G:G").HorizontalAlignment = xlCenter
    ws.Columns("I:I").HorizontalAlignment = xlCenter
    ws.Columns("J:J").HorizontalAlignment = xlCenter
    ws.Columns("L:L").HorizontalAlignment = xlCenter
    
    MsgBox "¡Fixture generado con números separados, Fran!", vbInformation
End Sub
