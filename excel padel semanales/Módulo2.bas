Attribute VB_Name = "Módulo2"

Sub ActualizarPuntos()
    Dim respuesta As VbMsgBoxResult
    Dim i As Long, j As Long, k As Long, ultimoJugador As Long
    Dim nombreJugador As String, celdaSuperior As String
    Dim puntosPartido As Variant
    
    ' --- SEGURIDAD ---
    respuesta = MsgBox("¿Deseas actualizar los puntos en la HOJA ACTUAL?" & vbCrLf & _
                "Asegúrate de estar en la fecha correcta.", vbQuestion + vbYesNo, "Confirmar Actualización")
    If respuesta = vbNo Then Exit Sub
    
    ' Forzar el cálculo de las fórmulas antes de procesar los datos
    Application.Calculate
    
    ' --- VARIABLES DE CONTROL DE FILAS ---
    Dim colJugador As Long
    Dim filaJugador As Long
    Dim filaRival As Long
    Dim encontrado As Boolean
    
    ' 1. PROCESO DE CÁLCULO (Escaneo fila por fila)
    For i = 22 To 33
        nombreJugador = Trim(Cells(i, 3).Value)
        If nombreJugador <> "" Then
            encontrado = False
            colJugador = 0
            
            ' Recorremos las filas reales de los partidos (5, 7, 9, 11, 13, 15, 17, 19)
            For k = 5 To 19 Step 2
                ' Comprobamos si el nombre está en la Columna C o E usando InStr
                If InStr(1, Cells(k, 3).Value, nombreJugador, vbTextCompare) > 0 Then
                    filaJugador = k
                    colJugador = 3
                    encontrado = True
                    Exit For
                ElseIf InStr(1, Cells(k, 5).Value, nombreJugador, vbTextCompare) > 0 Then
                    filaJugador = k
                    colJugador = 5
                    encontrado = True
                    Exit For
                End If
            Next k
            
            ' Si el jugador fue encontrado en la parte superior
            If encontrado Then
                celdaSuperior = Trim(Cells(filaJugador, colJugador).Value)

                ' --- DETECCIÓN DE FALTA ---
                If Right(UCase(celdaSuperior), 2) = " F" Or UCase(celdaSuperior) = "F" Then
                    If Cells(i, 8).Value = "" Then Cells(i, 8).Value = 0
                    If Cells(i, 10).Value = "" Then Cells(i, 10).Value = 0
                    Cells(i, 11).Value = Cells(i, 11).Value + 12 ' Puntos Perdidos por falta
                    Cells(i, 9).Value = Cells(i, 9).Value + 2    ' Sets Perdidos por falta
                Else
                    ' --- CASO JUGADOR PRESENTE ---
                    puntosPartido = Cells(filaJugador, 11).Value ' Columna K (Puntos de Partido)
                    
                    ' Aseguramos que la tabla de posiciones de abajo tenga números base (0)
                    If Cells(i, 8).Value = "" Then Cells(i, 8).Value = 0   ' Sets Ganados
                    If Cells(i, 9).Value = "" Then Cells(i, 9).Value = 0   ' Sets Perdidos
                    If Cells(i, 10).Value = "" Then Cells(i, 10).Value = 0 ' Puntos Ganados
                    If Cells(i, 11).Value = "" Then Cells(i, 11).Value = 0 ' Puntos Perdidos

                    ' --- CRUCE DE FILAS RIVALES ---
                    ' Parejas de Arriba: 5, 11, 17 -> El rival está abajo (+2)
                    ' Parejas de Abajo:  7, 13, 19 -> El rival está arriba (-2)
                    If filaJugador = 5 Or filaJugador = 11 Or filaJugador = 17 Then
                        filaRival = filaJugador + 2
                    Else
                        filaRival = filaJugador - 2
                    End If

                    Dim miSet1 As Variant, miSet2 As Variant
                    Dim rivalSet1 As Variant, rivalSet2 As Variant
                    Dim setsGanadosHoy As Integer, setsPerdidosHoy As Integer
                    
                    ' Lectura de los sets jugados por cada pareja
                    miSet1 = Cells(filaJugador, 6).Value      ' Mi 1º Set (F)
                    miSet2 = Cells(filaJugador, 7).Value      ' Mi 2º Set (G)
                    rivalSet1 = Cells(filaRival, 6).Value     ' 1º Set del Rival (F)
                    rivalSet2 = Cells(filaRival, 7).Value     ' G del Rival (G)
                    
                    setsGanadosHoy = 0
                    setsPerdidosHoy = 0
                    
                    ' Evaluar el Primer Set
                    If miSet1 <> "" And rivalSet1 <> "" Then
                        If Val(miSet1) > Val(rivalSet1) Then
                            setsGanadosHoy = setsGanadosHoy + 1
                        Else
                            setsPerdidosHoy = setsPerdidosHoy + 1
                        End If
                    End If
                    
                    ' Evaluar el Segundo Set
                    If miSet2 <> "" And rivalSet2 <> "" Then
                        If Val(miSet2) > Val(rivalSet2) Then
                            setsGanadosHoy = setsGanadosHoy + 1
                        Else
                            setsPerdidosHoy = setsPerdidosHoy + 1
                        End If
                    End If
                    
                    ' --- RESOLUCIÓN DEL SÚPER TIE-BREAK INVISIBLE (1-1) ---
                    If setsGanadosHoy = 1 And setsPerdidosHoy = 1 Then
                        If puntosPartido = 2 Then
                            setsGanadosHoy = 2
                            setsPerdidosHoy = 1
                        ElseIf puntosPartido = 1 Then
                            setsGanadosHoy = 1
                            setsPerdidosHoy = 2
                        End If
                    End If

                    ' Guardamos los sets acumulados en la tabla de posiciones de abajo
                    Cells(i, 8).Value = Cells(i, 8).Value + setsGanadosHoy
                    Cells(i, 9).Value = Cells(i, 9).Value + setsPerdidosHoy
                    
                    ' Sumamos los Puntos (games) ganados (H) y perdidos (I) de mi propia fila
                    Cells(i, 10).Value = Cells(i, 10).Value + Cells(filaJugador, 8).Value
                    Cells(i, 11).Value = Cells(i, 11).Value + Cells(filaJugador, 9).Value
                End If
            End If
        Else
            ' Limpiamos celdas si la fila de la tabla está vacía
            Range(Cells(i, 2), Cells(i, 13)).ClearContents
        End If
    Next i

    ' 2. AGRUPAR JUGADORES HACIA ARRIBA
    For i = 22 To 32
        If Trim(Cells(i, 3).Value) = "" Then
            For j = i + 1 To 33
                If Trim(Cells(j, 3).Value) <> "" Then
                    Range(Cells(j, 3), Cells(j, 13)).Copy Destination:=Cells(i, 3)
                    Range(Cells(j, 2), Cells(j, 13)).ClearContents
                    Exit For
                End If
            Next j
        End If
    Next i

    ' 3. DETERMINAR RANGO REAL Y ORDENAR
    ultimoJugador = 21
    For i = 22 To 33
        If Trim(Cells(i, 3).Value) <> "" Then ultimoJugador = i
    Next i

    If ultimoJugador >= 22 Then
        ActiveSheet.Sort.SortFields.Clear
        Range("B22:M" & ultimoJugador).Sort _
            Key1:=Range("L22:L" & ultimoJugador), Order1:=xlDescending, _
            Key2:=Range("I22:I" & ultimoJugador), Order2:=xlAscending, _
            Key3:=Range("M22:M" & ultimoJugador), Order3:=xlDescending, _
            Header:=xlNo
    End If

    ' 4. RENUMERACIÓN FINAL DE POSICIONES
    Dim puesto As Integer
    puesto = 1
    For i = 22 To 33
        If Trim(Cells(i, 3).Value) <> "" Then
            Cells(i, 2).Value = puesto
            puesto = puesto + 1
        Else
            Range(Cells(i, 2), Cells(i, 13)).ClearContents
        End If
    Next i

    ' Mensaje de aviso final
    MsgBox "¡Proceso terminado con éxito!" & vbCrLf & _
           "Hoja actual: " & ActiveSheet.Name & vbCrLf & _
           "Jugadores procesados: " & (puesto - 1), vbInformation, "Charvis System"
End Sub
