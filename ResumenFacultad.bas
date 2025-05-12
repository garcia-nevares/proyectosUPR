Attribute VB_Name = "ResumenFacultad"
Option Explicit

Sub GenerarResumenPorFacultad()
    Dim wsResumen As Worksheet, wsFac As Worksheet
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim lastRow As Long, lastCol As Long, destRow As Long
    Dim colFAC As Long, colMatr As Long, colCupo As Long, colCUPO_MIN As Long
    Dim colElearn As Long, colTipo As Long, colAccion As Long, colTipoSecc As Long
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
        .Range("A1:H1").Value = Array("FACULTAD", "CURS_SECC", "TIPO_DE_SECCION", "ELEARN", "ACTIVIDAD", _
                                      "CUPO", "CUPO_MINIMO", "MATR")
        destRow = 2
    End With

    ' Consolidar datos de todas las hojas de facultad
    For Each wsFac In wb.Sheets
        If wsFac.Name <> "Cursos" And wsFac.Name <> "Resumen" Then
            With wsFac
                lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row

                colFAC = .Rows(1).Find("FAC", , xlValues, xlWhole).Column
                colMatr = .Rows(1).Find("MATR", , xlValues, xlWhole).Column
                colCupo = .Rows(1).Find("CUPO MAX", , xlValues, xlWhole).Column
                colCUPO_MIN = .Rows(1).Find("CUPO MIN", , xlValues, xlWhole).Column
                colElearn = .Rows(1).Find("Modalidad", , xlValues, xlWhole).Column
                colAccion = .Rows(1).Find("Act", , xlValues, xlWhole).Column
                colTipoSecc = .Rows(1).Find("TIPO SECC", , xlValues, xlWhole).Column

                For i = 2 To lastRow
                    If Trim(.Cells(i, colCupo).Value) <> "" Then
                        With wsResumen
                            .Cells(destRow, 1).Value = wsFac.Name
                            .Cells(destRow, 2).Value = wsFac.Cells(i, colFAC).Value
                            .Cells(destRow, 3).Value = wsFac.Cells(i, colTipoSecc).Value
                            .Cells(destRow, 4).Value = wsFac.Cells(i, colElearn).Value
                            .Cells(destRow, 5).Value = wsFac.Cells(i, colAccion).Value
                            .Cells(destRow, 6).Value = wsFac.Cells(i, colCupo).Value
                            .Cells(destRow, 7).Value = wsFac.Cells(i, colCUPO_MIN).Value
                            .Cells(destRow, 8).Value = wsFac.Cells(i, colMatr).Value
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
        .Cells(1, 9).Value = "POR_DEBAJO_MIN"
        .Cells(2, 9).Resize(lastRow - 1).FormulaR1C1 = "=IF(RC8 < RC7,1,0)"
        
        .Cells(1, 10).Value = "CUPOMAX_ES_CUPO"
        .Cells(2, 10).Resize(lastRow - 1).FormulaR1C1 = "=IF(RC6 = RC7,1,0)"
        
        .Cells(1, 11).Value = "MATR_RESTR"
        .Cells(2, 11).Resize(lastRow - 1).FormulaR1C1 = "=IF(RC6=0,1,0)"
        
        .Cells(1, 12).Value = "SOBRECUPOS"
        .Cells(2, 12).Resize(lastRow - 1).FormulaR1C1 = "=IF(RC8 > RC6,1,0)"
    End With

    ' Crear tabla dinámica
    Dim ptCache As PivotCache, pt As PivotTable
    Dim ptRange As Range
    lastCol = wsResumen.Cells(1, wsResumen.Columns.Count).End(xlToLeft).Column
    Set tblRng = wsResumen.Range(wsResumen.Cells(1, 1), wsResumen.Cells(lastRow, lastCol))

    Set ptCache = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=tblRng)
    Set pt = wsResumen.PivotTables.Add(PivotCache:=ptCache, TableDestination:=wsResumen.Cells(2, 15), TableName:="ResumenPT")

    With pt
        .ClearAllFilters
        .PivotFields("FACULTAD").Orientation = xlRowField
        .PivotFields("ACTIVIDAD").Orientation = xlRowField
        
        .PivotFields("TIPO_DE_SECCION").Orientation = xlPageField
        .PivotFields("ELEARN").Orientation = xlPageField
        
        .AddDataField .PivotFields("POR_DEBAJO_MIN"), "Por debajo cupo mínimo", xlSum
        .AddDataField .PivotFields("CUPOMAX_ES_CUPO"), "Cupo max = cupo mínimo", xlSum
        .AddDataField .PivotFields("MATR_RESTR"), "Matr. restringida (Cupo=0)", xlSum
        .AddDataField .PivotFields("SOBRECUPOS"), "En sobrecupo", xlSum
    End With

    wsResumen.Columns.AutoFit
    Application.ScreenUpdating = True

    MsgBox "Resumen generado correctamente.", vbInformation
End Sub


