Option Explicit

Sub ProcesarCursosTxt()
    '*****************************************************
    ' Macro: ProcesarCursosTxt
    ' Descripción:
    '  - Importa un archivo .txt de cursos delimitado por "|".
    '  - Limpia y transforma los datos.
    '  - Lee mínimos de cupo desde la tabla Cupos_Minimos.
    '  - Añade filtros automáticos.
    '  - Divide los cursos en hojas por facultad (según tabla Facultades).
    '  - Ajusta columnas: mueve FAC al inicio, esconde columnas específicas.
    '  - Guarda el archivo como nuevo workbook.
    '*****************************************************
    
    Dim sPaso As String
    Dim tmp As Long
    On Error GoTo ErrHandler
    
    ' 1) Seleccionar archivo TXT
    sPaso = "Seleccionar archivo TXT"
    Dim txtFilePath As String
    txtFilePath = Application.GetOpenFilename( _
        "Archivos de texto (*.txt),*.txt", , _
        "Seleccione el archivo TXT con los cursos")
    If txtFilePath = "False" Then
        MsgBox "Cancelado", vbInformation
        Exit Sub
    End If
    
    ' 2) Crear nuevo libro y hoja "Cursos"
    sPaso = "Crear nuevo libro y hoja"
    Dim wbNew As Workbook
    Set wbNew = Workbooks.Add(xlWBATWorksheet)
    Dim wsNew As Worksheet
    Set wsNew = wbNew.Worksheets(1)
    wsNew.Name = "Cursos"
    
    ' 3) Importar datos del TXT
    sPaso = "Importar datos TXT"
    With wsNew.QueryTables.Add( _
        Connection:="TEXT;" & txtFilePath, _
        Destination:=wsNew.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileOtherDelimiter = "|"
        .TextFileColumnDataTypes = Array(xlGeneralFormat)
        .AdjustColumnWidth = True
        .Refresh BackgroundQuery:=False
    End With
    
    ' 4) Verificar columnas requeridas según TablaColumnas
    sPaso = "Verificar columnas requeridas según TablaColumnas"
    
    Dim loCols As ListObject
    Dim colRequeridas As Collection
    Dim falta As Collection, sobra As Collection
    Dim headerCelda As Range
    Dim colItem As Variant
    Dim wsColSrc As Worksheet
    Dim dictHeaders As Object, dictRequeridas As Object
    Dim i As Long
    
    Set colRequeridas = New Collection
    Set falta = New Collection
    Set sobra = New Collection
    Set dictHeaders = CreateObject("Scripting.Dictionary")
    Set dictRequeridas = CreateObject("Scripting.Dictionary")
    
    ' Buscar la tabla TablaColumnas
    For Each wsColSrc In ThisWorkbook.Worksheets
        On Error Resume Next
        Set loCols = wsColSrc.ListObjects("TablaColumnas")
        On Error GoTo 0
        If Not loCols Is Nothing Then Exit For
    Next wsColSrc
    
    If loCols Is Nothing Then
        MsgBox "No se encontró la tabla 'TablaColumnas'.", vbCritical
        Exit Sub
    End If
    
    ' Leer columnas esperadas desde ColDTAA, ignorando las que tengan OrdenArchivoDTAA = 9999
    Dim ordenVal As Variant
    For i = 1 To loCols.ListRows.Count
        colItem = Trim(loCols.ListRows(i).Range(1, 1).Value)
        ordenVal = loCols.ListRows(i).Range(1, 2).Value ' Segunda columna = OrdenArchivoDTAA

        If Len(colItem) > 0 And Not IsEmpty(ordenVal) And ordenVal <> 9999 Then
            colRequeridas.Add colItem
            dictRequeridas(UCase(colItem)) = True
        End If
    Next i
    
    ' Construir diccionario con los encabezados actuales del TXT
    For Each headerCelda In wsNew.Range("A1").CurrentRegion.Rows(1).Cells
        colItem = UCase(Trim(headerCelda.Value))
        If Len(colItem) > 0 Then dictHeaders(colItem) = True
    Next headerCelda
    
    ' Verificar columnas faltantes
    For Each colItem In colRequeridas
        If Not dictHeaders.Exists(UCase(colItem)) Then
            falta.Add colItem
        End If
    Next colItem
    
    ' Verificar columnas inesperadas (sobrantes)
    For Each colItem In dictHeaders.Keys
        If Not dictRequeridas.Exists(colItem) Then
            sobra.Add colItem
        End If
    Next colItem
    
    ' Mostrar mensaje si hay diferencias
    If falta.Count > 0 Or sobra.Count > 0 Then
        Dim mensaje As String
        mensaje = "Resultado de la validación de columnas:" & vbNewLine & vbNewLine
        
        If falta.Count > 0 Then
            mensaje = mensaje & "FALTAN estas columnas en el archivo TXT:" & vbNewLine
            For Each colItem In falta
                mensaje = mensaje & "   - " & colItem & vbNewLine
            Next colItem
            mensaje = mensaje & vbNewLine
        End If
        
        If sobra.Count > 0 Then
            mensaje = mensaje & " NO SE ESPERABAN estas columnas en el archivo TXT:" & vbNewLine
            For Each colItem In sobra
                mensaje = mensaje & "   - " & colItem & vbNewLine
            Next colItem
            mensaje = mensaje & vbNewLine
        End If
    
        MsgBox mensaje, vbCritical, "Validación de columnas"
        If falta.Count > 0 Then Exit Sub
    End If
  
    ' 5) Eliminar filas con DEPT=ESUP
    sPaso = "Eliminar filas con DEPT=ESUP"
    Dim lastRow As Long
    lastRow = wsNew.Cells(wsNew.Rows.Count, "A").End(xlUp).Row
    Dim j As Long
    Dim colDept As Long
    colDept = wsNew.Rows(1).Find("DEPT", , xlValues, xlWhole).Column
    For i = lastRow To 2 Step -1
        If UCase(wsNew.Cells(i, colDept).Value) = "ESUP" Then
            wsNew.Rows(i).Delete
        End If
    Next i
    

    ' 6) Formatear CURS_SECC como XXXX9999-yyy
    sPaso = "Formatear CURS_SECC como XXXX9999-yyy"

    Dim colCurs As Long
    colCurs = wsNew.Rows(1).Find("CURS_SECC", , xlValues, xlWhole).Column

    For i = 2 To lastRow
        Dim valRaw As String
        valRaw = Trim(wsNew.Cells(i, colCurs).Value)
        
        If Len(valRaw) = 11 Then
            wsNew.Cells(i, colCurs).Value = _
                Left(valRaw, 8) & "-" & Right(valRaw, 3)
        End If
    Next i

    ' 7) Consolidar columnas EDIF y SALON en EDIF-SALON
    sPaso = "Consolidar columnas EDIF y SALON"

    Dim colEDIF As Long, colSALON As Long, colEDIF_SALON As Long

    colEDIF = wsNew.Rows(1).Find("EDIF", , xlValues, xlWhole).Column
    colSALON = wsNew.Rows(1).Find("SALON", , xlValues, xlWhole).Column

    ' Insertar columna nueva al final
    colEDIF_SALON = wsNew.Cells(1, wsNew.Columns.Count).End(xlToLeft).Column + 1
    wsNew.Cells(1, colEDIF_SALON).Value = "EDIF-SALON"

    ' Llenar valores combinados
    For i = 2 To lastRow
        Dim edifVal As String, salonVal As String
        edifVal = Trim(wsNew.Cells(i, colEDIF).Value)
        salonVal = Trim(wsNew.Cells(i, colSALON).Value)

        If edifVal <> "" And salonVal <> "" Then
            wsNew.Cells(i, colEDIF_SALON).Value = edifVal & "-" & salonVal
        ElseIf edifVal <> "" Then
            wsNew.Cells(i, colEDIF_SALON).Value = edifVal
        ElseIf salonVal <> "" Then
            wsNew.Cells(i, colEDIF_SALON).Value = salonVal
        Else
            wsNew.Cells(i, colEDIF_SALON).Value = ""
        End If
    Next i

    ' Eliminar columnas originales (de derecha a izquierda para evitar desplazamiento)
    If colEDIF > colSALON Then
        wsNew.Columns(colEDIF).Delete
        wsNew.Columns(colSALON).Delete
    Else
        wsNew.Columns(colSALON).Delete
        wsNew.Columns(colEDIF).Delete
    End If

    ' 8) Consolidar horarios por CURS_SECC
    sPaso = "Consolidar múltiples filas por CURS_SECC"

    Dim colDays As Long, colHin As Long, colHout As Long
    Dim colHorario As Long, colCursoSecc As Long
    Dim dictSecciones As Object, key As String
    Dim rowData As Variant
    Dim hInStr As String, hOutStr As String

    Set dictSecciones = CreateObject("Scripting.Dictionary")

    colCursoSecc = wsNew.Rows(1).Find("CURS_SECC", , xlValues, xlWhole).Column
    colSALON = wsNew.Rows(1).Find("EDIF-SALON", , xlValues, xlWhole).Column
    colDays = wsNew.Rows(1).Find("DAYS", , xlValues, xlWhole).Column
    colHin = wsNew.Rows(1).Find("H-IN", , xlValues, xlWhole).Column
    colHout = wsNew.Rows(1).Find("H-OUT", , xlValues, xlWhole).Column

    ' Crear columna nueva para HORARIO
    colHorario = wsNew.Cells(1, wsNew.Columns.Count).End(xlToLeft).Column + 1
    wsNew.Cells(1, colHorario).Value = "HORARIO"

    lastRow = wsNew.Cells(wsNew.Rows.Count, 1).End(xlUp).Row

    ' Recorrer y agrupar por CURS_SECC
    For i = 2 To lastRow
        key = Trim(wsNew.Cells(i, colCursoSecc).Value)
        
        If IsDate(wsNew.Cells(i, colHin).Value) Then
            hInStr = Format(wsNew.Cells(i, colHin).Value, "hh:mm")
        Else
            hInStr = wsNew.Cells(i, colHin).Text
        End If
        
        If IsDate(wsNew.Cells(i, colHout).Value) Then
            hOutStr = Format(wsNew.Cells(i, colHout).Value, "hh:mm")
        Else
            hOutStr = wsNew.Cells(i, colHout).Text
        End If

        If Not dictSecciones.Exists(key) Then
            dictSecciones.Add key, _
                Array(wsNew.Cells(i, colSALON).Value, _
                      wsNew.Cells(i, colDays).Value, _
                      hInStr & "-" & hOutStr)
        Else
            rowData = dictSecciones(key)
            rowData(0) = rowData(0) & vbLf & wsNew.Cells(i, colSALON).Value
            rowData(1) = rowData(1) & vbLf & wsNew.Cells(i, colDays).Value
            rowData(2) = rowData(2) & vbLf & hInStr & "-" & hOutStr
            dictSecciones(key) = rowData
        End If
    Next i

    ' Eliminar todas las filas menos una por CURS_SECC
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    For i = lastRow To 2 Step -1
        key = Trim(wsNew.Cells(i, colCursoSecc).Value)
        If seen.Exists(key) Then
            wsNew.Rows(i).Delete
        Else
            seen.Add key, True
        End If
    Next i

    ' Recalcular último row
    lastRow = wsNew.Cells(wsNew.Rows.Count, 1).End(xlUp).Row

    ' Insertar los valores consolidados
    For i = 2 To lastRow
        key = Trim(wsNew.Cells(i, colCursoSecc).Value)
        If dictSecciones.Exists(key) Then
            rowData = dictSecciones(key)
            wsNew.Cells(i, colSALON).Value = rowData(0)
            wsNew.Cells(i, colDays).Value = rowData(1)
            wsNew.Cells(i, colHorario).Value = rowData(2)
        End If
    Next i
    
    ' Eliminar columnas H-IN y H-OUT (de derecha a izquierda para evitar desplazamiento)
    colHin = wsNew.Rows(1).Find("H-IN", , xlValues, xlWhole).Column
    colHout = wsNew.Rows(1).Find("H-OUT", , xlValues, xlWhole).Column
    If colHin > colHout Then
        wsNew.Columns(colHin).Delete
        wsNew.Columns(colHout).Delete
    Else
        wsNew.Columns(colHout).Delete
        wsNew.Columns(colHin).Delete
    End If

    ' 9) Consolidar datos de PROFx y LOD%x en una sola celda
    sPaso = "Consolidar columnas PROFx y LOD%x"
    Dim colProf(1 To 6) As Long, colLOD(1 To 6) As Long, colID(1 To 6) As Long
    Dim celda As Range
    Dim colList As New Collection
    
    ' Buscar posiciones de columnas PROF, LOD%, ID
    For i = 1 To 6
        Set celda = wsNew.Rows(1).Find(What:="PROF" & i, LookAt:=xlPart)
        If Not celda Is Nothing Then colProf(i) = celda.Column: colList.Add colProf(i)
        
        Set celda = wsNew.Rows(1).Find(What:="LOD%" & i, LookAt:=xlPart)
        If Not celda Is Nothing Then colLOD(i) = celda.Column: colList.Add colLOD(i)
        
        Set celda = wsNew.Rows(1).Find(What:="ID" & i, LookAt:=xlPart)
        If Not celda Is Nothing Then colID(i) = celda.Column: colList.Add colID(i)
    Next i
    
    ' Crear columna de consolidación
    Dim colConsol As Long
    colConsol = wsNew.Cells(1, wsNew.Columns.Count).End(xlToLeft).Column + 1
    wsNew.Cells(1, colConsol).Value = "DETALLE_PROF_LOD"
    
    lastRow = wsNew.Cells(wsNew.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        Dim sDetalle As String: sDetalle = ""
        For j = 1 To 6
        Dim sProf As String
        Dim sID As String, last4 As String, sLOD As String
        sProf = ""
        If colID(j) > 0 Then sID = Trim(wsNew.Cells(i, colID(j)).Value) Else sID = ""

        If colProf(j) > 0 Then sProf = Trim(wsNew.Cells(i, colProf(j)).Value)

        If sID <> "" And sID <> "0" Then
            If sProf = "" Then
                If Len(sID) >= 4 Then
                    last4 = Right(sID, 4)
                Else
                    last4 = sID
                End If
                sProf = "XXX-XX-" & last4
            End If
        End If

        If sProf <> "" Then
            If sDetalle <> "" Then sDetalle = sDetalle & vbNewLine
            sDetalle = sDetalle & sProf

            sLOD = ""
            If colLOD(j) > 0 Then sLOD = Trim(wsNew.Cells(i, colLOD(j)).Value)
            If sLOD <> "" Then sDetalle = sDetalle & " (" & sLOD & "%)"
        End If
        Next j
        wsNew.Cells(i, colConsol).Value = sDetalle
    Next i
    
    ' Eliminar columnas PROF, LOD%, ID
    Dim vCols() As Long
    ReDim vCols(1 To colList.Count)
    For j = 1 To colList.Count
        vCols(j) = colList(j)
    Next j
    
    Dim swapped As Boolean
    Do
        swapped = False
        For j = LBound(vCols) To UBound(vCols) - 1
            If vCols(j) < vCols(j + 1) Then
                tmp = vCols(j): vCols(j) = vCols(j + 1): vCols(j + 1) = tmp
                swapped = True
            End If
        Next j
    Loop While swapped
    
    For j = 1 To UBound(vCols)
        wsNew.Columns(vCols(j)).Delete
    Next j
    
    ' 10) Añadir columna TIPO_DE_SECCION
    sPaso = "Añadir columna TIPO_DE_SECCION"

    Dim colTipoDeSeccion As Long
    colTipoDeSeccion = wsNew.Cells(1, wsNew.Columns.Count).End(xlToLeft).Column + 1
    wsNew.Cells(1, colTipoDeSeccion).Value = "TIPO_DE_SECCION"

    Dim dictTipoClave As Object
    Set dictTipoClave = CreateObject("Scripting.Dictionary")

    Dim colCursSecc As Long
    colCursSecc = wsNew.Rows(1).Find("CURS_SECC", , xlValues, xlWhole).Column
    Dim colTipoSecc As Long
    colTipoSecc = wsNew.Rows(1).Find("TIPO", , xlValues, xlWhole).Column

    ' Contar ocurrencias por clave (primeros 8 de CURS_SECC + tipo)
    For i = 2 To lastRow
        Dim claveTipo As String
        claveTipo = Left(wsNew.Cells(i, colCursSecc).Value, 8) & "|" & Trim(wsNew.Cells(i, colTipoSecc).Value)
        dictTipoClave(claveTipo) = dictTipoClave(claveTipo) + 1
    Next i

    ' Asignar M o U según conteo
    For i = 2 To lastRow
        claveTipo = Left(wsNew.Cells(i, colCursSecc).Value, 8) & "|" & Trim(wsNew.Cells(i, colTipoSecc).Value)
        wsNew.Cells(i, colTipoDeSeccion).Value = IIf(dictTipoClave(claveTipo) > 1, "M", "U")
    Next i

    ' 11) Determinar NIVEL, ELEARN, CUPO_MINIMO y %_AL_CUPO_MIN
    sPaso = "Determinar nivel, modalidad y estado de cupo"
    Dim wsSrc As Worksheet, lo As ListObject
    Dim colNivel As Long, colElearn As Long, colMatr As Long, colTipo As Long
    Dim colCupo As Long, colEstado As Long, colCupoMinimo As Long, colPOR As Long
    Dim nivelKey As String, elearnKey As String, tipodeseccionKey As String
    Dim tipoCurso As String, clave As String
    Dim hdrMatch As Variant
    Dim minCupo As Double
    Dim lr As ListRow
    
    ' Buscar tabla Cupos_Minimos
    For Each wsSrc In ThisWorkbook.Worksheets
        Set lo = Nothing
        On Error Resume Next
        Set lo = wsSrc.ListObjects("Cupos_Minimos")
        On Error GoTo ErrHandler
        If Not lo Is Nothing Then Exit For
    Next wsSrc
    If lo Is Nothing Then
        MsgBox "No se encontró la tabla 'Cupos_Minimos'.", vbCritical
        Exit Sub
    End If

    ' Insertar columnas CUPO_MINIMO y %_AL_CUPO_MIN
    colCupo = wsNew.Rows(1).Find("CUPO", , xlValues, xlWhole).Column
    colMatr = wsNew.Rows(1).Find("MATR", , xlValues, xlWhole).Column
    
    wsNew.Columns(colCupo).Insert Shift:=xlToRight
    colCupoMinimo = colCupo
    wsNew.Cells(1, colCupoMinimo).Value = "CUPO_MINIMO"
    
    colMatr = wsNew.Rows(1).Find("MATR", , xlValues, xlWhole).Column
    wsNew.Columns(colMatr + 1).Insert Shift:=xlToRight
    colEstado = colMatr + 1
    wsNew.Cells(1, colEstado).Value = "%_AL_CUPO_MIN"

    ' Recalcular posiciones de columnas necesarias
    colNivel = wsNew.Rows(1).Find("NIVEL", , xlValues, xlWhole).Column
    colTipo = wsNew.Rows(1).Find("TIPO", , xlValues, xlWhole).Column
    colElearn = wsNew.Rows(1).Find("ELEARN", , xlValues, xlWhole).Column
    colTipoDeSeccion = wsNew.Rows(1).Find("TIPO_DE_SECCION", , xlValues, xlWhole).Column
    colCupoMinimo = wsNew.Rows(1).Find("CUPO_MINIMO", , xlValues, xlWhole).Column
    colMatr = wsNew.Rows(1).Find("MATR", , xlValues, xlWhole).Column
    
    ' Buscar cupo mínimo y calcular estado
    lastRow = wsNew.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    Dim colTipoTbl As Long
    colTipoTbl = Application.Match("Tipo", lo.HeaderRowRange, 0)
    
    For i = 2 To lastRow
        nivelKey = IIf(UCase(wsNew.Cells(i, colNivel).Value) = "SS", "S", "G")
        elearnKey = IIf(Trim(wsNew.Cells(i, colElearn).Value) = "" _
                        Or UCase(wsNew.Cells(i, colElearn).Value) = "P", _
                        "P", "NP")
        tipodeseccionKey = wsNew.Cells(i, colTipoDeSeccion).Value
        clave = nivelKey & "-" & elearnKey & "-" & tipodeseccionKey
        tipoCurso = Trim(wsNew.Cells(i, colTipo).Value)

        minCupo = 0
        hdrMatch = Application.Match(clave, lo.HeaderRowRange, 0)
        If IsError(hdrMatch) Then
            hdrMatch = Application.Match(clave & "-O", lo.HeaderRowRange, 0)
        End If
        
        If Not IsError(hdrMatch) Then
            For Each lr In lo.ListRows
                If Trim(UCase(lr.Range(1, colTipoTbl).Value)) = Trim(UCase(tipoCurso)) Then
                    minCupo = lr.Range(1, hdrMatch).Value
                    Exit For
                End If
            Next lr
        End If
        
        If IsEmpty(minCupo) Or IsError(minCupo) Then minCupo = 0
        
        wsNew.Cells(i, colCupoMinimo).Value = minCupo
        
		If minCupo = 0 Then
			wsNew.Cells(i, colEstado).Value = "Min 0"
        ElseIf wsNew.Cells(i, colMatr).Value >= minCupo Then
            wsNew.Cells(i, colEstado).Value = "Ok"
        Else
            wsNew.Cells(i, colEstado).Value = _
                Format((minCupo - wsNew.Cells(i, colMatr).Value) / minCupo, "0%")
        End If
    Next i

    ' 12) Ajustar formato de columnas principales (se hace más adelante, eliminado)
    ' wsNew.Columns.AutoFit

    ' 13) Mover columna TIPO_DE_SECCION después de CURS_SECC
    Dim colAfterCursSecc As Long
    colAfterCursSecc = wsNew.Rows(1).Find("CURS_SECC", , xlValues, xlWhole).Column + 1
    colTipoDeSeccion = wsNew.Rows(1).Find("TIPO_DE_SECCION", , xlValues, xlWhole).Column
    If colTipoDeSeccion <> colAfterCursSecc Then
        wsNew.Columns(colTipoDeSeccion).Cut
        wsNew.Columns(colAfterCursSecc).Insert Shift:=xlToRight
        Application.CutCopyMode = False
    End If

    ' 14) Formatear columna CUPO_MINIMO como General
    colCupoMinimo = wsNew.Rows(1).Find("CUPO_MINIMO", , xlValues, xlWhole).Column
    wsNew.Columns(colCupoMinimo).NumberFormat = "General"

    ' 15) Aplicar formato condicional en %_AL_CUPO_MIN
    Dim rngEstado As Range
    colEstado = wsNew.Rows(1).Find("%_AL_CUPO_MIN", , xlValues, xlWhole).Column
    Set rngEstado = wsNew.Range(wsNew.Cells(2, colEstado), wsNew.Cells(lastRow, colEstado))
    
    rngEstado.FormatConditions.Delete
    With rngEstado.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""OK""")
        .Interior.Color = RGB(146, 208, 80) ' Verde para OK
        .Font.Color = RGB(0, 97, 0)
    End With
    With rngEstado.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Min 0""")
        .Interior.Color = RGB(255, 153, 0) ' Anaranjado
        .Font.Color = RGB(0, 97, 0)
    End With
    With rngEstado.FormatConditions.AddDatabar
        .MinPoint.Modify newtype:=xlConditionValueLowestValue
        .MaxPoint.Modify newtype:=xlConditionValueHighestValue
        .BarColor.Color = RGB(255, 0, 0) ' Rojo
        .BarFillType = xlDataBarFillGradient
        .NegativeBarFormat.ColorType = xlDataBarColor
        .ShowValue = True
        .Direction = xlRTL
    End With
    
  
    ' Recalcular posiciones de columnas tras inserciones y movimientos
    colCupo = wsNew.Rows(1).Find("CUPO", , xlValues, xlWhole).Column
    colMatr = wsNew.Rows(1).Find("MATR", , xlValues, xlWhole).Column
    colCupoMinimo = wsNew.Rows(1).Find("CUPO_MINIMO", , xlValues, xlWhole).Column

    ' 17) Renombrar y recalcular columna POR% ? %_AL_CUPO_MAX_SOBREC
    sPaso = "Recalcular %_AL_CUPO_MAX_SOBREC"
    
    ' Recalcular posiciones por si han cambiado
    colCupo = wsNew.Rows(1).Find("CUPO", , xlValues, xlWhole).Column
    colMatr = wsNew.Rows(1).Find("MATR", , xlValues, xlWhole).Column
    colCupoMinimo = wsNew.Rows(1).Find("CUPO_MINIMO", , xlValues, xlWhole).Column
    colPOR = wsNew.Rows(1).Find("POR%", , xlValues, xlWhole).Column
    
    wsNew.Cells(1, colPOR).Value = "%_AL_CUPO_MAX_SOBREC"
    
    Dim valCupo As Variant, valMatr As Variant, valCupoMin As Variant
    Dim cellValue As String
    
    For i = 2 To lastRow
        valCupo = wsNew.Cells(i, colCupo).Value
        valMatr = wsNew.Cells(i, colMatr).Value
        valCupoMin = wsNew.Cells(i, colCupoMinimo).Value
        
        If IsNumeric(valCupo) And IsNumeric(valMatr) And IsNumeric(valCupoMin) Then
            If valCupo = 0 Then
                If valCupoMin > 0 Then
                    cellValue = "MatRestr: " & Format(valMatr / valCupoMin, "0%")
                Else
                    cellValue = "MatRestr: N/A"
                End If
            ElseIf valMatr <= valCupo Then
                cellValue = Format(valMatr / valCupo, "0%")
            Else
                cellValue = "SobreC: " & Format(1 - (valMatr / valCupo), "0%")
            End If
        Else
            cellValue = "N/D"
        End If
		If valCupoMin = valCupo Then
			cellValue = "Cupo Min=Cupo Max. " & cellValue
		End If
        wsNew.Cells(i, colPOR).Value = cellValue
    Next i

    ' Formato condicional en %_AL_CUPO_MAX_SOBREC
    Dim rngPOR As Range
    Set rngPOR = wsNew.Range(wsNew.Cells(2, colPOR), wsNew.Cells(lastRow, colPOR))
    
    rngPOR.FormatConditions.Delete

    ' Amarillo si contiene "MatRestr" o "SobreC"
    With rngPOR.FormatConditions.Add(Type:=xlTextString, String:="MatRestr", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 255, 0)
        .Font.Color = RGB(156, 101, 0)
        .StopIfTrue = False
    End With
    With rngPOR.FormatConditions.Add(Type:=xlTextString, String:="SobreC", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 204, 204)
        .Font.Color = RGB(156, 101, 0)
        .StopIfTrue = False
    End With

    ' Databar verde solo si es numérico
    With rngPOR.FormatConditions.AddDatabar
        .MinPoint.Modify newtype:=xlConditionValueLowestValue
        .MaxPoint.Modify newtype:=xlConditionValueHighestValue
        .BarColor.Color = RGB(146, 208, 80)
        .BarFillType = xlDataBarFillGradient
        .NegativeBarFormat.ColorType = xlDataBarColor
        .ShowValue = True
    End With

    ' 18) Reorganizar, renombrar y describir columnas según TablaColumnas
    sPaso = "Aplicar estructura según TablaColumnas"
    
    Dim loTC As ListObject
    Dim colOrden As Variant, colNombre As Variant, colDescripcion As Variant
    Dim colIncluir As Variant, colEtiqueta As Variant
    Dim wsTC As Worksheet
    Dim fila As ListRow
    Dim colEncontrada As Range
    Dim nuevaColIndex As Long
    
    ' Buscar la tabla TablaColumnas
    Set loTC = Nothing
    For Each wsTC In ThisWorkbook.Worksheets
        On Error Resume Next
        Set loTC = wsTC.ListObjects("TablaColumnas")
        On Error GoTo ErrHandler
        If Not loTC Is Nothing Then Exit For
    Next wsTC
    
    If loTC Is Nothing Then
        MsgBox "No se encontró la tabla 'TablaColumnas'", vbCritical
        Exit Sub
    End If
    
    nuevaColIndex = 1
    
    For Each fila In loTC.ListRows
        colOrden = fila.Range(1, loTC.ListColumns("OrdenInforme").Index).Value
        colNombre = Trim(fila.Range(1, loTC.ListColumns("ColDTAA").Index).Value)
        colEtiqueta = Trim(fila.Range(1, loTC.ListColumns("ColumnaInforme").Index).Value)
        colDescripcion = Trim(fila.Range(1, loTC.ListColumns("Descripción").Index).Value)
        colIncluir = UCase(Trim(fila.Range(1, loTC.ListColumns("Incluir").Index).Value))
    
        If colOrden <> 9999 And colNombre <> "" Then
            Set colEncontrada = Nothing
            Set colEncontrada = wsNew.Rows(1).Find(What:=colNombre, LookAt:=xlWhole)
    
            If Not colEncontrada Is Nothing Then
                ' Mover la columna si es necesario
                If colEncontrada.Column <> nuevaColIndex Then
                    colEncontrada.EntireColumn.Cut
                    wsNew.Columns(nuevaColIndex).Insert Shift:=xlToRight
                    Application.CutCopyMode = False
                End If
    
                ' Renombrar si ColumnaInforme no está vacía
                If colEtiqueta <> "" Then
                    wsNew.Cells(1, nuevaColIndex).Value = colEtiqueta
                End If
    
                ' Agregar mensaje de entrada con la descripción
                With wsNew.Cells(1, nuevaColIndex).Validation
                    .Delete
                    .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
                    .InputTitle = "Descripción"
                    .InputMessage = colDescripcion
                    .IgnoreBlank = True
                    .InCellDropdown = False
                End With
    
                ' Eliminar si Incluir = "N"
                If colIncluir = "N" Then
                    wsNew.Columns(nuevaColIndex).Delete
                    GoTo SiguienteColumna ' Pasar al siguiente sin incrementar índice
                End If

                nuevaColIndex = nuevaColIndex + 1
            End If
        End If
SiguienteColumna:
    Next fila

    ' 19) Dividir en hojas según Facultades
    sPaso = "Dividir datos en hojas según Facultades"
    Dim loFacultades As ListObject
    Dim dictHojas As Object, dictHojasNombres As Object
    Dim filtros As Variant
    Dim hojaNombre As String
    Dim colFAC As Long
    Dim nuevaWs As Worksheet
    Dim copiarRango As Range
    Dim colFACnueva As Long, colHidden As Long
    Dim nombreFacultad As String
    Dim facValor As Variant, agruparValor As Variant
    Dim colName As Variant
    
    Set dictHojas = CreateObject("Scripting.Dictionary")
    Set dictHojasNombres = CreateObject("Scripting.Dictionary")
    
    ' Buscar tabla Facultades
    For Each wsSrc In ThisWorkbook.Worksheets
        Set loFacultades = Nothing
        On Error Resume Next
        Set loFacultades = wsSrc.ListObjects("Facultades")
        On Error GoTo ErrHandler
        If Not loFacultades Is Nothing Then Exit For
    Next wsSrc
    If loFacultades Is Nothing Then
        MsgBox "No se encontró la tabla 'Facultades'.", vbCritical
        Exit Sub
    End If
    
    ' Construir diccionarios
    For i = 1 To loFacultades.ListRows.Count
        facValor = Trim(loFacultades.ListRows(i).Range(1, 1).Value) ' FAC
        nombreFacultad = Trim(loFacultades.ListRows(i).Range(1, 2).Value) ' Nombre
        agruparValor = Trim(loFacultades.ListRows(i).Range(1, 3).Value) ' Agrupar
        
        If agruparValor = "N" Or agruparValor = "" Then
            dictHojas.Add facValor, Array(facValor)
            dictHojasNombres.Add facValor, nombreFacultad
        Else
            If Not dictHojas.Exists(agruparValor) Then
                dictHojas.Add agruparValor, Array(facValor)
                dictHojasNombres.Add agruparValor, nombreFacultad
            Else
                Dim tempArray As Variant
                tempArray = dictHojas(agruparValor)
                ReDim Preserve tempArray(UBound(tempArray) + 1)
                tempArray(UBound(tempArray)) = facValor
                dictHojas(agruparValor) = tempArray
            End If
        End If
    Next i
    
    ' Identificar columna FAC en hoja Cursos
    colFAC = wsNew.Rows(1).Find("FAC", , xlValues, xlWhole).Column
        
    ' Alinear verticalmente arriba todas las celdas en la hoja principal
    wsNew.Cells.VerticalAlignment = xlTop

    ' Añadir columnas de decisión antes de dividir por facultades
    sPaso = "Añadir columnas de validación y comentario"

    Dim colInsertBase As Long
    Dim colAccion As Long, colJustif As Long, colFinal As Long, colComent As Long
    colInsertBase = wsNew.Cells(1, wsNew.Columns.Count).End(xlToLeft).Column + 1

    ' 1. ACCIÓN
    colAccion = colInsertBase
    wsNew.Cells(1, colAccion).Value = "ACCIÓN"
    With wsNew.Range(wsNew.Cells(2, colAccion), wsNew.Cells(wsNew.Rows.Count, colAccion))
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="A,C"
        .Validation.IgnoreBlank = True
        .Validation.InCellDropdown = True
    End With
    With wsNew.Cells(1, colAccion).Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
        .InputTitle = "ACCIÓN"
        .InputMessage = "A = dejar abierta; C = sección cerrada"
    End With

    ' 2. JUSTIFICACIÓN
    colJustif = colAccion + 1
    wsNew.Cells(1, colJustif).Value = "JUSTIFICACIÓN"
    With wsNew.Cells(1, colJustif).Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
        .InputTitle = "JUSTIFICACIÓN"
        .InputMessage = "Explicación para dejar abierta o cerrar la sección"
    End With

    ' 3. ACCIÓN FINAL
    colFinal = colJustif + 1
    wsNew.Cells(1, colFinal).Value = "ACCIÓN FINAL"
    With wsNew.Range(wsNew.Cells(2, colFinal), wsNew.Cells(wsNew.Rows.Count, colFinal))
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="A,C,D"
        .Validation.IgnoreBlank = True
        .Validation.InCellDropdown = True
    End With
    With wsNew.Cells(1, colFinal).Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
        .InputTitle = "ACCIÓN FINAL"
        .InputMessage = "A = dejar abierta; C = cerrada por Facultad; D = cerrada por DAA"
    End With

    ' 4. COMENTARIO DAA
    colComent = colFinal + 1
    wsNew.Cells(1, colComent).Value = "COMENTARIO DAA"
    With wsNew.Cells(1, colComent).Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
        .InputTitle = "COMENTARIO DAA"
        .InputMessage = "Observaciones o instrucciones del DAA"
    End With

    ' Aplicar filtros automáticos a la hoja principal
    sPaso = "Aplicar filtros automáticos"
    Dim dataRange As Range
    Set dataRange = wsNew.Range("A1").CurrentRegion
    dataRange.AutoFilter

	' Congelar fila 1
	wsNew.Range("A2").Select
	wsNew.Application.ActiveWindow.FreezePanes = True

    
    sPaso = "Dividir datos en hojas según Facultades"

     
     ' Crear hojas nuevas
    For Each facValor In dictHojas.Keys
        filtros = dictHojas(facValor)
        
        ' Nombre de hoja basado en el campo Nombre
        hojaNombre = dictHojasNombres(facValor)
        hojaNombre = Application.WorksheetFunction.Substitute(hojaNombre, ":", "-")
        hojaNombre = Application.WorksheetFunction.Substitute(hojaNombre, "/", "-")
        hojaNombre = Application.WorksheetFunction.Substitute(hojaNombre, "\", "-")
        hojaNombre = Left(hojaNombre, 31)
        
        ' Crear nueva hoja
        Set nuevaWs = wbNew.Sheets.Add(After:=wbNew.Sheets(wbNew.Sheets.Count))
        nuevaWs.Name = hojaNombre
        
        ' Copiar encabezados
        dataRange.Rows(1).Copy Destination:=nuevaWs.Rows(1)
        
        ' Filtrar y copiar datos
        wsNew.AutoFilterMode = False
        dataRange.AutoFilter Field:=colFAC, Criteria1:=filtros, Operator:=xlFilterValues
        
        On Error Resume Next
        Set copiarRango = dataRange.Offset(1, 0).Resize(dataRange.Rows.Count - 1, dataRange.Columns.Count).SpecialCells(xlCellTypeVisible)
        On Error GoTo ErrHandler
        
        If Not copiarRango Is Nothing Then
            copiarRango.Copy Destination:=nuevaWs.Cells(2, 1)
        End If
        Set copiarRango = Nothing
        
        ' Ajustes en cada hoja
                
        ' Aplicar autofiltros
        nuevaWs.Range("A1").CurrentRegion.AutoFilter

		' Proteger hojas permitiendo modificar solo esas 4 columnas
		Dim cell As Range
		nuevaWs.Unprotect
		nuevaWs.Cells.Locked = True
		
		lastRow = nuevaWs.Cells(nuevaWs.Rows.Count, 1).End(xlUp).Row
		
		Dim colA As Long, colJ As Long, colF As Long, colC As Long
		Dim colPerc As Long
		
		colA = nuevaWs.Rows(1).Find("ACCIÓN", , xlValues, xlWhole).Column
		colJ = nuevaWs.Rows(1).Find("JUSTIFICACIÓN", , xlValues, xlWhole).Column
		' colF = nuevaWs.Rows(1).Find("ACCIÓN FINAL", , xlValues, xlWhole).Column
		' colC = nuevaWs.Rows(1).Find("COMENTARIO DAA", , xlValues, xlWhole).Column
		
		Range(nuevaWs.Cells(2, colA), nuevaWs.Cells(lastRow, colA)).Locked = False
		Range(nuevaWs.Cells(2, colJ), nuevaWs.Cells(lastRow, colJ)).Locked = False
		' Range(nuevaWs.Cells(2, colF), nuevaWs.Cells(lastRow, colF)).Locked = False
		' Range(nuevaWs.Cells(2, colC), nuevaWs.Cells(lastRow, colC)).Locked = False

		colPerc = nuevaWs.Rows(1).Find("Perc a min.", , xlValues, xlWhole).Column
		


		' Establecer formato visual y restringir selección solo a celdas desbloqueadas
		With nuevaWs
			' Aplicar color a celdas desbloqueadas
			.Range(.Cells(2, colA), .Cells(lastRow, colA)).Interior.Color = RGB(255, 255, 153) ' Amarillo claro
			.Range(.Cells(2, colJ), .Cells(lastRow, colJ)).Interior.Color = RGB(255, 255, 153)
			'  .Range(.Cells(2, colF), .Cells(lastRow, colF)).Interior.Color = RGB(255, 255, 153)
			'  .Range(.Cells(2, colC), .Cells(lastRow, colC)).Interior.Color = RGB(255, 255, 153)

			' Ajustar ancho de columnas para que se vea todo y filas
			.Columns.AutoFit
			.Rows.AutoFit
			.Columns(colJ).ColumnWidth = 40 ' Justificación: more room for text

			' Aplicar filtro: ocultar filas con "Ok"
			.Range("A1").CurrentRegion.AutoFilter Field:=colPerc, Criteria1:="<>Ok"
			' Congelar fila 1
			.Range("A2").Select
			.Application.ActiveWindow.FreezePanes = True

			' Progeger la hoja
			' .EnableSelection = xlUnlockedCells
			.Protect Password:="facultad2025", AllowFiltering:=True, _
				AllowInsertingRows:=False, AllowDeletingRows:=False, _
				AllowInsertingColumns:=False, AllowDeletingColumns:=False, _
				AllowSorting:=True, AllowFormattingCells:=False, _
				AllowFormattingColumns:=False, AllowFormattingRows:=False
		End With

    Next facValor
    
    ' Ajustar ancho de columnas para que se vea todo y filas
    wsNew.Columns.AutoFit
    wsNew.Rows.AutoFit
    
    ' Limpiar criterios de filtro en hoja principal antes de guardar
    If wsNew.FilterMode Then wsNew.ShowAllData

    ' Progeger la hoja
    ' wsNew.Protect Password:="facultad2025", AllowFiltering:=True, _
    '     AllowInsertingRows:=False, AllowDeletingRows:=False, _
    '     AllowInsertingColumns:=False, AllowDeletingColumns:=False, _
    '     AllowSorting:=False


    ' 20) Guardar el nuevo archivo
    sPaso = "Guardar nuevo archivo"
    Dim savePath As Variant
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="CuposProcesados.xlsx", _
        FileFilter:="Archivos de Excel (*.xlsx), *.xlsx")
    If savePath <> False Then
        wbNew.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
        MsgBox "Archivo guardado en:" & vbNewLine & savePath, vbInformation, "Proceso completado"
    End If

    Exit Sub

' Manejo de errores
ErrHandler:
    MsgBox "Se produjo un error en el paso: " & sPaso & vbNewLine & _
           "Descripción: " & Err.Description, _
           vbCritical, "Error en macro"
End Sub

