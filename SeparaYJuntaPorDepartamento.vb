Sub DividirArchivoPorDepartamento()
    Dim wbOriginal As Workbook, wbNuevo As Workbook
    Dim wsOriginal As Worksheet, wsNuevo As Worksheet
    Dim ultimaFila As Long, lastCol As Long
    Dim departamentos As Collection, fila As Long, dept As Variant
    Dim rutaBase As String, nombreArchivo As String, rutaCompleta As String
    Dim colAccion As Long, colJustificacion As Long, i As Long
    Dim fd As FileDialog
    Dim nuevaFila As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wbOriginal = ThisWorkbook
    Set wsOriginal = wbOriginal.Sheets(1)

    ' Preguntar carpeta de salida
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Seleccione la carpeta donde se guardarán los archivos por departamento"
    If fd.Show <> -1 Then
        MsgBox "Operación cancelada.", vbExclamation
        Exit Sub
    End If
    rutaBase = fd.SelectedItems(1) & Application.PathSeparator

    ' Identificar columnas de ACCIÓN y JUSTIFICACIÓN
    lastCol = wsOriginal.Cells(1, wsOriginal.Columns.Count).End(xlToLeft).Column
    For i = 1 To lastCol
        Select Case Trim(UCase(wsOriginal.Cells(1, i).Value))
            Case "ACCIÓN": colAccion = i
            Case "JUSTIFICACIÓN": colJustificacion = i
        End Select
    Next i
    If colAccion = 0 Or colJustificacion = 0 Then
        MsgBox "No se encontraron las columnas ACCIÓN y JUSTIFICACIÓN", vbCritical
        Exit Sub
    End If

    ' Crear lista única de departamentos visibles
    Set departamentos = New Collection
    On Error Resume Next
    For fila = 2 To wsOriginal.Cells(wsOriginal.Rows.Count, "C").End(xlUp).Row
        If Not wsOriginal.Rows(fila).Hidden Then
            Dim d As String
            d = Trim(wsOriginal.Cells(fila, "C").Value)
            If d <> "" Then departamentos.Add d, d
        End If
    Next fila
    On Error GoTo 0

    ' Generar archivos
    For Each dept In departamentos
        Set wbNuevo = Workbooks.Add
        Set wsNuevo = wbNuevo.Sheets(1)
        wsNuevo.Name = "Datos"

        ' Copiar encabezado
        wsOriginal.Range(wsOriginal.Cells(1, 1), wsOriginal.Cells(1, lastCol)).Copy Destination:=wsNuevo.Range("A1")
        nuevaFila = 2

        ' Copiar solo filas visibles y del departamento correspondiente
        ultimaFila = wsOriginal.Cells(wsOriginal.Rows.Count, "C").End(xlUp).Row
        For fila = 2 To ultimaFila
            If Not wsOriginal.Rows(fila).Hidden Then
                If Trim(wsOriginal.Cells(fila, "C").Value) = dept Then
                    wsOriginal.Range(wsOriginal.Cells(fila, 1), wsOriginal.Cells(fila, lastCol)).Copy
                    wsNuevo.Cells(nuevaFila, 1).PasteSpecial Paste:=xlPasteAll
                    nuevaFila = nuevaFila + 1
                End If
            End If
        Next fila
        Application.CutCopyMode = False

        ' Ajustar JUSTIFICACIÓN
        With wsNuevo.Columns(colJustificacion)
            .ColumnWidth = 60
            .WrapText = True
        End With
        
        ' Habilitar autofiltro
        wsNuevo.Range("A1").AutoFilter
        
        ' Ajustar columnas
        wsNuevo.Cells.EntireColumn.AutoFit

        ' Proteger hoja con permisos para ordenar y filtrar
        wsNuevo.Cells.Locked = True
        wsNuevo.Rows(1).Locked = True
        wsNuevo.Columns(colAccion).Locked = False
        wsNuevo.Columns(colJustificacion).Locked = False
        
        wsNuevo.Protect Password:="departamento", AllowFiltering:=True, _
                AllowInsertingRows:=False, AllowDeletingRows:=False, _
                AllowInsertingColumns:=False, AllowDeletingColumns:=False, _
                AllowSorting:=True, AllowFormattingCells:=False, _
                AllowFormattingColumns:=False, AllowFormattingRows:=False

        ' Guardar archivo (eliminando si ya existe)
        nombreArchivo = "C51-HU-" & dept & ".xlsx"
        rutaCompleta = rutaBase & nombreArchivo
        If Dir(rutaCompleta) <> "" Then Kill rutaCompleta
        wbNuevo.SaveAs rutaCompleta
        wbNuevo.Close SaveChanges:=False
    Next dept

    MsgBox "Archivos por departamento guardados correctamente.", vbInformation
    Application.ScreenUpdating = True
End Sub

Sub ConsolidarRespuestasDepartamentos()
    Dim carpeta As String, archivo As String
    Dim wbDestino As Workbook, wsDestino As Worksheet
    Dim wbFuente As Workbook, wsFuente As Worksheet
    Dim filaFuente As Long, filaDestino As Long
    Dim colCurso As Long, colAccion As Long, colJustif As Long
    Dim cursoValor As Variant, cursoBuscar As Range
    Dim i As Long
    Dim logWS As Worksheet
    Dim filaLogEncontrados As Long, filaLogNoEncontrados As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wbDestino = ThisWorkbook
    Set wsDestino = wbDestino.Sheets(1) ' Asegúrate de que esta sea tu hoja de datos

    ' Columnas en archivo original
    colCurso = 4 ' Columna D = Curso

    ' Detectar columnas ACCIÓN y JUSTIFICACIÓN en destino
    For i = 1 To wsDestino.Cells(1, wsDestino.Columns.Count).End(xlToLeft).Column
        Select Case Trim(UCase(wsDestino.Cells(1, i).Value))
            Case "ACCIÓN": colAccion = i
            Case "JUSTIFICACIÓN": colJustif = i
        End Select
    Next i
    If colAccion = 0 Or colJustif = 0 Then
        MsgBox "No se encontraron las columnas 'ACCIÓN' y 'JUSTIFICACIÓN' en el archivo original.", vbCritical
        Exit Sub
    End If

    ' Crear hoja de log (eliminar si existe)
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("LOG_CONSOLIDACIÓN").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set logWS = Worksheets.Add(After:=wbDestino.Sheets(wbDestino.Sheets.Count))
    logWS.Name = "LOG_CONSOLIDACIÓN"
    logWS.Range("A1").Value = "Cursos encontrados y actualizados"
    logWS.Range("B1").Value = "Cursos NO encontrados"
    filaLogEncontrados = 2
    filaLogNoEncontrados = 2

    ' Seleccionar carpeta
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Selecciona la carpeta con los archivos de departamentos completados"
        If .Show <> -1 Then Exit Sub
        carpeta = .SelectedItems(1) & Application.PathSeparator
    End With

    ' Procesar archivos
    archivo = Dir(carpeta & "C51-HU-*.xlsx")
    Do While archivo <> ""
        Set wbFuente = Workbooks.Open(carpeta & archivo, ReadOnly:=True)
        On Error Resume Next
        Set wsFuente = wbFuente.Sheets("Datos")
        On Error GoTo 0

        If Not wsFuente Is Nothing Then
            For filaFuente = 2 To wsFuente.Cells(wsFuente.Rows.Count, colCurso).End(xlUp).Row
                cursoValor = Trim(wsFuente.Cells(filaFuente, colCurso).Value)
                If Len(cursoValor) > 0 Then
                    Set cursoBuscar = wsDestino.Columns(colCurso).Find(What:=cursoValor, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not cursoBuscar Is Nothing Then
                        filaDestino = cursoBuscar.Row
                        wsDestino.Cells(filaDestino, colAccion).Value = wsFuente.Cells(filaFuente, colAccion).Value
                        wsDestino.Cells(filaDestino, colJustif).Value = wsFuente.Cells(filaFuente, colJustif).Value
                        logWS.Cells(filaLogEncontrados, 1).Value = cursoValor
                        filaLogEncontrados = filaLogEncontrados + 1
                    Else
                        logWS.Cells(filaLogNoEncontrados, 2).Value = cursoValor
                        filaLogNoEncontrados = filaLogNoEncontrados + 1
                    End If
                End If
            Next filaFuente
        End If

        wbFuente.Close SaveChanges:=False
        Set wsFuente = Nothing
        archivo = Dir
    Loop

    MsgBox "Consolidación completada. Revisa la hoja 'LOG_CONSOLIDACIÓN' para verificar resultados.", vbInformation
    Application.ScreenUpdating = True
End Sub


