Attribute VB_Name = "ModuloHorarios"
Sub GenerarJSON()
    ' ============================================================
    ' Genera horarios.json a partir de la hoja "Datos" (tabla plana
    ' Dia | Sesion | Docente | Tipo | Grupo | Materia | Aula | Actividad)
    ' en vez de leer texto libre de las 5 hojas de dias.
    ' ============================================================

    Dim wsDatos As Worksheet, wsDocentes As Worksheet
    On Error Resume Next
    Set wsDatos = ThisWorkbook.Sheets("Datos")
    Set wsDocentes = ThisWorkbook.Sheets("Docentes")
    On Error GoTo 0
    If wsDatos Is Nothing Then
        MsgBox "No se encuentra la hoja 'Datos'. Genera primero la tabla plana.", vbExclamation, "Falta hoja Datos"
        Exit Sub
    End If
    If wsDocentes Is Nothing Then
        MsgBox "No se encuentra la hoja 'Docentes' (columnas Código | Docente).", vbExclamation, "Falta hoja Docentes"
        Exit Sub
    End If

    ' --- Cargar tabla de codigos reales: Nombre -> Codigo (col. A=Codigo, col. B=Nombre) ---
    Dim codigoPorNombre As Object
    Set codigoPorNombre = CreateObject("Scripting.Dictionary")
    Dim filaDoc As Long, ultimaFilaDoc As Long
    ultimaFilaDoc = wsDocentes.Cells(wsDocentes.Rows.Count, 2).End(xlUp).Row
    For filaDoc = 2 To ultimaFilaDoc
        Dim codReal As String, nomReal As String
        codReal = Trim(wsDocentes.Cells(filaDoc, 1).Value)
        nomReal = Trim(wsDocentes.Cells(filaDoc, 2).Value)
        If nomReal <> "" And codReal <> "" Then
            If Not codigoPorNombre.Exists(nomReal) Then
                codigoPorNombre.Add nomReal, codReal
            End If
        End If
    Next filaDoc

    Dim nombresSinCodigo As String
    nombresSinCodigo = ""

    ' --- Metadatos de sesiones (igual que el JSON original) ---
    Dim sesiones() As String
    sesiones = Split("1,2,3,3R,4,5,6,7,8", ",")

    Dim sesLabels(8) As String
    sesLabels(0) = "Guardia manana"
    sesLabels(1) = "Sesion 1"
    sesLabels(2) = "Sesion 2"
    sesLabels(3) = "Sesion 3"
    sesLabels(4) = "Recreo"
    sesLabels(5) = "Sesion 4"
    sesLabels(6) = "Sesion 5"
    sesLabels(7) = "Sesion 6"
    sesLabels(8) = "Sesion 7"

    Dim sesHoras(8) As String
    sesHoras(0) = "08:00-08:30"
    sesHoras(1) = "08:30-09:20"
    sesHoras(2) = "09:25-10:15"
    sesHoras(3) = "10:20-11:10"
    sesHoras(4) = "11:10-11:35"
    sesHoras(5) = "11:35-12:25"
    sesHoras(6) = "12:30-13:25"
    sesHoras(7) = "13:30-14:20"
    sesHoras(8) = "14:25-15:15"

    ' --- Traduccion de la etiqueta de sesion (columna B de Datos) a la clave JSON ---
    Dim mapaSesion As Object
    Set mapaSesion = CreateObject("Scripting.Dictionary")
    mapaSesion.Add "Guardia mañana", "1"
    mapaSesion.Add "Sesión 1", "2"
    mapaSesion.Add "Sesión 2", "3"
    mapaSesion.Add "Sesión 3", "3R"
    mapaSesion.Add "Recreo", "4"
    mapaSesion.Add "Sesión 4", "5"
    mapaSesion.Add "Sesión 5", "6"
    mapaSesion.Add "Sesión 6", "7"
    mapaSesion.Add "Sesión 7", "8"

    ' --- Recopilar todos los docentes a partir de la hoja Datos ---
    Dim docentes As Object
    Set docentes = CreateObject("Scripting.Dictionary")

    Dim ultimaFila As Long
    ultimaFila = wsDatos.Cells(wsDatos.Rows.Count, 1).End(xlUp).Row

    Dim fila As Long
    For fila = 2 To ultimaFila
        Dim diaNombre As String, sesionEtiqueta As String, nombreDocente As String
        Dim tipo As String, grupo As String, materia As String, aula As String, actividad As String

        diaNombre = Trim(wsDatos.Cells(fila, 1).Value)
        If diaNombre = "" Then GoTo SiguienteFila

        sesionEtiqueta = Trim(wsDatos.Cells(fila, 2).Value)
        nombreDocente = Trim(wsDatos.Cells(fila, 3).Value)
        tipo = Trim(wsDatos.Cells(fila, 4).Value)
        grupo = Trim(wsDatos.Cells(fila, 5).Value)
        materia = Trim(wsDatos.Cells(fila, 6).Value)
        aula = Trim(wsDatos.Cells(fila, 7).Value)
        actividad = Trim(wsDatos.Cells(fila, 8).Value)

        ' Traducir tipo "Clase"/"Otro" (hoja Datos) -> "clase"/"actividad" (JSON)
        Dim tipoJSON As String
        If LCase(tipo) = "clase" Then
            tipoJSON = "clase"
        Else
            tipoJSON = "actividad"
        End If

        ' Traducir etiqueta de sesion a clave JSON
        Dim sesKey As String
        If mapaSesion.Exists(sesionEtiqueta) Then
            sesKey = mapaSesion(sesionEtiqueta)
        Else
            sesKey = sesionEtiqueta ' fallback, no deberia ocurrir
        End If

        If Not docentes.Exists(nombreDocente) Then
            docentes.Add nombreDocente, CreateObject("Scripting.Dictionary")
        End If
        Dim docenteData As Object
        Set docenteData = docentes(nombreDocente)

        If Not docenteData.Exists(diaNombre) Then
            docenteData.Add diaNombre, CreateObject("Scripting.Dictionary")
        End If
        Dim diaData As Object
        Set diaData = docenteData(diaNombre)

        ' Si una sesion ya tiene actividades (combo de varias siglas en la misma celda
        ' original), se concatenan con espacio para mantener compatibilidad con el
        ' formato anterior (p.ej. "G CHL").
        Dim registro As Object
        If diaData.Exists(sesKey) Then
            Set registro = diaData(sesKey)
            If tipoJSON = "actividad" And registro("tipo") = "actividad" Then
                registro("sigla") = Trim(registro("sigla") & " " & actividad)
            End If
        Else
            Set registro = CreateObject("Scripting.Dictionary")
            registro.Add "tipo", tipoJSON
            registro.Add "sigla", actividad
            registro.Add "grupo", grupo
            registro.Add "materia", materia
            registro.Add "aula", aula
            diaData.Add sesKey, registro
        End If

SiguienteFila:
    Next fila

    ' --- Construir JSON ---
    Dim json As String
    json = "{" & Chr(10)

    json = json & "  ""sesiones"": {" & Chr(10)
    Dim s As Integer
    For s = 0 To 8
        json = json & "    """ & sesiones(s) & """: {""label"": """ & sesLabels(s) & """, ""hora"": """ & sesHoras(s) & """}"
        If s < 8 Then json = json & ","
        json = json & Chr(10)
    Next s
    json = json & "  }," & Chr(10)

    json = json & "  ""docentes"": {" & Chr(10)

    Dim keys() As Variant
    keys = docentes.Keys

    Dim i As Integer, j As Integer, tmp As Variant
    For i = 0 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If keys(i) > keys(j) Then
                tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
            End If
        Next j
    Next i

    For i = 0 To UBound(keys)
        Dim nombreKey As String
        nombreKey = keys(i)
        Dim docenteObj As Object
        Set docenteObj = docentes(nombreKey)

        Dim codigo As String
        If codigoPorNombre.Exists(nombreKey) Then
            codigo = codigoPorNombre(nombreKey)
        Else
            codigo = GenerarCodigo(nombreKey) ' fallback si el docente no esta en la hoja Docentes
            nombresSinCodigo = nombresSinCodigo & "- " & nombreKey & Chr(10)
        End If

        json = json & "    """ & codigo & """: {" & Chr(10)
        json = json & "      ""nombre"": """ & EscapeJSON(nombreKey) & """," & Chr(10)
        json = json & "      ""horario"": {" & Chr(10)

        Dim diaKeys() As Variant
        diaKeys = docenteObj.Keys

        Dim dk As Integer
        For dk = 0 To UBound(diaKeys)
            Dim diaKey As String
            diaKey = diaKeys(dk)
            Dim diaObj As Object
            Set diaObj = docenteObj(diaKey)

            json = json & "        """ & diaKey & """: {" & Chr(10)

            Dim sesKeys() As Variant
            sesKeys = diaObj.Keys

            Dim sk As Integer
            For sk = 0 To UBound(sesKeys)
                Dim sesKeyOut As String
                sesKeyOut = sesKeys(sk)
                Dim reg As Object
                Set reg = diaObj(sesKeyOut)

                json = json & "          """ & sesKeyOut & """: {"
                json = json & """tipo"": """ & reg("tipo") & """, "
                json = json & """sigla"": """ & EscapeJSON(reg("sigla")) & """, "
                json = json & """grupo"": """ & EscapeJSON(reg("grupo")) & """, "
                json = json & """materia"": """ & EscapeJSON(reg("materia")) & """, "
                json = json & """aula"": """ & EscapeJSON(reg("aula")) & """}"

                If sk < UBound(sesKeys) Then json = json & ","
                json = json & Chr(10)
            Next sk

            json = json & "        }"
            If dk < UBound(diaKeys) Then json = json & ","
            json = json & Chr(10)
        Next dk

        json = json & "      }" & Chr(10)
        json = json & "    }"
        If i < UBound(keys) Then json = json & ","
        json = json & Chr(10)
    Next i

    json = json & "  }" & Chr(10)
    json = json & "}"

    Dim rutaExcel As String
    rutaExcel = ThisWorkbook.Path
    Dim rutaJSON As String
    rutaJSON = rutaExcel & "\horarios.json"

    Dim fileNum As Integer
    fileNum = FreeFile
    Open rutaJSON For Output As #fileNum
    Print #fileNum, json
    Close #fileNum

    If nombresSinCodigo <> "" Then
        MsgBox "Aviso: los siguientes docentes de la hoja 'Datos' no se encontraron en la hoja 'Docentes' y se les generó un código automático (revisar nombres):" & Chr(10) & Chr(10) & nombresSinCodigo, vbExclamation, "Docentes sin código"
    End If

    MsgBox "horarios.json generado correctamente desde la hoja Datos en:" & Chr(10) & rutaJSON, vbInformation, "Exportacion completada"
End Sub

Function GenerarCodigo(nombre As String) As String
    Dim partes() As String
    partes = Split(nombre, " ")
    Dim codigo As String
    codigo = ""
    Dim p As Integer
    For p = 0 To UBound(partes)
        If Len(partes(p)) > 0 Then
            If partes(p) <> "DE" And partes(p) <> "DEL" And partes(p) <> "LA" And partes(p) <> "LOS" Then
                codigo = codigo & Left(UCase(partes(p)), 1)
            End If
        End If
        If Len(codigo) >= 6 Then Exit For
    Next p
    GenerarCodigo = Left(codigo & "XXXXXX", 6)
End Function

Function EscapeJSON(s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, Chr(10), "\n")
    s = Replace(s, Chr(13), "")
    EscapeJSON = s
End Function
