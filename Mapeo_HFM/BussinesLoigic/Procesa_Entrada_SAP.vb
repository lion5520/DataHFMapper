Imports System
Imports System.Data.SQLite
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel

Public Class Procesa_Entrada_SAP

    ''' <summary>
    ''' Lee un Excel (hoja "base") y vuelca su contenido en t_in_sap de SQLite.
    ''' Optimizado: lee todo el rango en un solo Array, luego itera en memoria.
    ''' </summary>
    ''' <param name="rutaExcel">Ruta al .xlsx de entrada</param>
    ''' <param name="rutaSqlite">Ruta al .sqlite donde insertar</param>
    Public Shared Sub Ejecutar(rutaExcel As String, rutaSqlite As String)
        Const sheetName As String = "Sheet1"
        Dim connString As String = $"Data Source={rutaSqlite};Version=3;"
        Dim conn As SQLiteConnection = Nothing
        Dim tx As SQLiteTransaction = Nothing

        Dim xlApp As Application = Nothing
        Dim wb As Workbook = Nothing
        Dim ws As Worksheet = Nothing
        Dim dataRange As Range = Nothing
        Dim allData As Object(,) = Nothing

        Try
            ' 1) Abrir conexión y empezar transacción
            conn = New SQLiteConnection(connString)
            conn.Open()
            tx = conn.BeginTransaction()

            ' 2) Limpiar tabla de destino
            Using cmdDel As New SQLiteCommand("DELETE FROM t_in_sap", conn, tx)
                cmdDel.ExecuteNonQuery()
            End Using

            ' 3) Limpia secuencia del ID de la tabla 
            Using cmdDel As New SQLiteCommand("DELETE FROM sqlite_sequence WHERE name='t_in_sap';", conn, tx)
                cmdDel.ExecuteNonQuery()
            End Using

            ' 4) Iniciar Excel y ubicar la hoja
            xlApp = New Application With {.Visible = False, .DisplayAlerts = False}
            wb = xlApp.Workbooks.Open(rutaExcel, ReadOnly:=True)

            ' Intentar hoja "base"
            Dim found As Boolean = False
            For Each sht As Worksheet In wb.Worksheets
                If String.Equals(sht.Name, sheetName, StringComparison.OrdinalIgnoreCase) Then
                    ws = sht
                    found = True
                    Exit For
                End If
            Next
            If Not found Then
                Throw New Exception($"No se encontró hoja '{sheetName}'.")
            End If

            ' 4) Detectar última fila con datos en columna A
            Dim lastRow As Integer = ws.Cells(ws.Rows.Count, 1).End(XlDirection.xlUp).Row
            If lastRow < 2 Then
                Return ' nada que procesar
            End If

            ' 5) Leer todo el rango de datos (desde fila 2 hasta lastRow, columnas A:AR)
            Const totalCols As Integer = 44
            dataRange = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, totalCols))
            allData = DirectCast(dataRange.Value2, Object(,))

            ' 6) Liberar la Range COM de inmediato
            Marshal.ReleaseComObject(dataRange)
            dataRange = Nothing

            ' 7) Preparar INSERT con parámetros (@p1...@p44)
            Dim cols As New List(Of String) From {
                "sociedad", "numero_cuenta", "texto_explicativo", "texto_cab_documento",
                "centro_de_coste", "clase_documento", "fecha_documento", "fecha_contabilizacion",
                "asignacion", "importe_ml", "saldo_acum", "importe_md",
                "sociedad_gl_asociada", "numero_documento", "periodo", "texto",
                "referencia", "centro_de_beneficio", "moneda", "orden",
                "ejercicio", "ej_mes", "deudor_acreedor", "nombre",
                "cta_contrapartida", "ceco_cebe", "deudor_acreedor_2", "agrup",
                "rpt", "des_rpt", "region", "tipo_cliente",
                "des_cliente", "depto", "des_depto", "linea_negocio",
                "des_linea_negocio", "cuenta_mayor_hfm", "descripcion_cuenta_sific",
                "cuenta_oracle", "descripcion_cuenta_oracle", "agrupador_tipo",
                "agrupador_cuenta", "agrupador_detalle"
            }
            Dim colList = String.Join(",", cols)
            Dim paramList = String.Join(",", Enumerable.Range(1, cols.Count).Select(Function(i) "@p" & i.ToString()))
            Dim insertSql = $"INSERT INTO t_in_sap ({colList}) VALUES ({paramList})"

            Using cmdIns As New SQLiteCommand(insertSql, conn, tx)
                ' Añadir y preparar parámetros
                For i As Integer = 1 To cols.Count
                    cmdIns.Parameters.Add(New SQLiteParameter($"@p{i}", DbType.String))
                Next
                cmdIns.Prepare()

                ' 8) Iterar en memoria y ejecutar INSERT
                Dim rowCount = allData.GetLength(0)
                Dim colCount = allData.GetLength(1)
                For r As Integer = 1 To rowCount
                    For c As Integer = 1 To colCount
                        Dim val = allData(r, c)
                        Dim prm = cmdIns.Parameters($"@p{c}")
                        ' Columnas 10-12 como REAL
                        If c >= 10 AndAlso c <= 12 Then
                            prm.Value = If(val IsNot Nothing, Convert.ToDouble(val), 0.0)
                        Else
                            prm.Value = If(val IsNot Nothing, val.ToString(), String.Empty)
                        End If
                    Next
                    cmdIns.ExecuteNonQuery()
                Next
            End Using

            ' 9) Commit
            tx.Commit()

        Catch ex As Exception
            ' Rollback si falla
            If tx IsNot Nothing Then
                Try : tx.Rollback() : Catch : End Try
            End If
            Throw

        Finally
            ' 10) Cierre limpio de Excel COM
            Try
                If wb IsNot Nothing Then wb.Close(False)
                If xlApp IsNot Nothing Then xlApp.Quit()
            Catch : End Try

            If ws IsNot Nothing Then Marshal.ReleaseComObject(ws)
            If wb IsNot Nothing Then Marshal.ReleaseComObject(wb)
            If xlApp IsNot Nothing Then Marshal.ReleaseComObject(xlApp)

            ws = Nothing
            wb = Nothing
            xlApp = Nothing

            GC.Collect()
            GC.WaitForPendingFinalizers()

            ' 11) Cerrar y disponer SQLite
            If conn IsNot Nothing Then
                conn.Close()
                conn.Dispose()
                conn = Nothing
            End If
        End Try
    End Sub

End Class
