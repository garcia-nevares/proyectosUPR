Option Explicit

Sub GenerarResumenPorFacultad()
    Dim wsResumen As Worksheet, wsFac As Worksheet
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim lastRow As Long, lastCol As Long, destRow As Long
    Dim colFAC As Long, colMatr As Long, colCupo As Long, colCUPO_MIN As Long
    Dim colModalidad As Long, colTipo As Long, colActividad As Long, colTipoSecc As Long
    Dim colDepart As Long, colCursSecc As Long
    Dim tblRng As Range
    Dim i As Long

    Application.ScreenUpdating = False

    ' Crear hoja de resumen
    On Error Resume Next: Set wsResumen = wb.Sheets("Resumen"): On Error GoTo 0
    If Not wsResumen Is Nothing Then Application.DisplayAlerts = False: wsResumen.Delete: Application.DisplayAlerts = True
    Set wsResumen = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsResumen.Name = "Resumen"

    ' Encabezados
    With wsResumen
        .Range("A1:I1").Value = Array("Facultad", "Depart", "Curso", "Tipo de sección", "Modalidad", "Actividad", _
                                      "Cupo mín", "Cupo max", "Matriculado")
        destRow = 2
    End With

    ' Consolidar datos de todas las hojas de facultad
    For Each wsFac In wb.Sheets
        If wsFac.Name <> "Cursos" And wsFac.Name <> "Resumen" Then
            With wsFac
                lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row

                colFAC = .Rows(1).Find("Fac", , xlValues, xlWhole).Column
                colCursSecc = .Rows(1).Find("Curso", , xlValues, xlWhole).Column
                colDepart = .Rows(1).Find("Depart", , xlValues, xlWhole).Column
                colTipoSecc = .Rows(1).Find("Tipo Secc", , xlValues, xlWhole).Column
                colModalidad = .Rows(1).Find("Modalidad", , xlValues, xlWhole).Column
                colActividad = .Rows(1).Find("Act", , xlValues, xlWhole).Column
                colCUPO_MIN = .Rows(1).Find("CUPO MIN", , xlValues, xlWhole).Column
                colCupo = .Rows(1).Find("CUPO MAX", , xlValues, xlWhole).Column
                colMatr = .Rows(1).Find("MATR", , xlValues, xlWhole).Column

                For i = 2 To lastRow
                    If Trim(.Cells(i, colCupo).Value) <> "" Then
                        With wsResumen
                            .Cells(destRow, 1).Value = wsFac.Name
                            .Cells(destRow, 2).Value = wsFac.Cells(i, colDepart).Value
                            .Cells(destRow, 3).Value = wsFac.Cells(i, colCursSecc).Value
                            .Cells(destRow, 4).Value = wsFac.Cells(i, colTipoSecc).Value
                            .Cells(destRow, 5).Value = wsFac.Cells(i, colModalidad).Value
                            .Cells(destRow, 6).Value = wsFac.Cells(i, colActividad).Value
                            .Cells(destRow, 7).Value = wsFac.Cells(i, colCUPO_MIN).Value
                            .Cells(destRow, 8).Value = wsFac.Cells(i, colCupo).Value
                            .Cells(destRow, 9).Value = wsFac.Cells(i, colMatr).Value
                        End With
                        destRow = destRow + 1
                    End If
                Next i
            End With
        End If
    Next wsFac

    ' Agregar columnas auxiliares con fórmulas
    With wsResumen
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        .Cells(1, 10).Value = "Bajo mínimo"
        .Cells(2, 10).Resize(lastRow - 1).FormulaR1C1 = "=IF(RC9 < RC7,1,0)"
        
        .Cells(1, 11).Value = "Min = Max"
        .Cells(2, 11).Resize(lastRow - 1).FormulaR1C1 = "=IF(RC7 = RC8,1,0)"
        
        .Cells(1, 12).Value = "Max < Min"
        .Cells(2, 12).Resize(lastRow - 1).FormulaR1C1 = "=IF(RC8 < RC7,1,0)"
        
        .Cells(1, 13).Value = "Mat restringida"
        .Cells(2, 13).Resize(lastRow - 1).FormulaR1C1 = "=IF(RC8=0,1,0)"
        
        .Cells(1, 14).Value = "Con sobrecupo"
        .Cells(2, 14).Resize(lastRow - 1).FormulaR1C1 = "=IF(RC9 > RC8,1,0)"
    End With

    ' Crear tabla dinámica
    Dim ptCache As PivotCache, pt As PivotTable
    Dim ptRange As Range
    lastCol = wsResumen.Cells(1, wsResumen.Columns.Count).End(xlToLeft).Column
        Set tblRng = wsResumen.Range(wsResumen.Cells(1, 1), wsResumen.Cells(lastRow, lastCol))

    Set ptCache = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=tblRng)
    Set pt = wsResumen.PivotTables.Add(PivotCache:=ptCache, TableDestination:=wsResumen.Cells(2, 18), TableName:="ResumenPT")

    With pt
        .ClearAllFilters
        .PivotFields("FACULTAD").Orientation = xlRowField
        .PivotFields("ACTIVIDAD").Orientation = xlRowField
        
        .PivotFields("Tipo de sección").Orientation = xlPageField
        .PivotFields("Modalidad").Orientation = xlPageField
        
        .AddDataField .PivotFields("Curso"), "Total secciones", xlCount
        .AddDataField .PivotFields("Bajo mínimo"), "Por debajo cupo mínimo", xlSum
        .AddDataField .PivotFields("Min = Max"), "Cupo max = cupo mínimo", xlSum
        .AddDataField .PivotFields("Max < Min"), "Cupo max < cupo min", xlSum
        .AddDataField .PivotFields("Mat restringida"), "Matr. restringida (Cupo=0)", xlSum
        .AddDataField .PivotFields("Con sobrecupo"), "En sobrecupo", xlSum
    End With

    wsResumen.Columns.AutoFit
    Application.ScreenUpdating = True

    MsgBox "Resumen generado correctamente.", vbInformation
End Sub



