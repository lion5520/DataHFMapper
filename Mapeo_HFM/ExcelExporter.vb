Imports System
Imports System.Data
Imports System.Data.SQLite
Imports System.Runtime.InteropServices
Imports Excel = Microsoft.Office.Interop.Excel
Imports Drawing = System.Drawing
Imports Forms = System.Windows.Forms

Public Class ExcelDbExporter

    ''' <summary>
    ''' Exporta a Excel el resultado de una consulta SQLite.
    ''' Puede recibir:
    '''   • tableName y whereClause (construye SELECT * FROM tableName [WHERE ...])  
    '''   • fullQuery (si no se proporcionan tableName ni whereClause).  
    ''' Oculta automáticamente cualquier columna llamada "id", 
    ''' formatea numéricos con dos decimales (moneda) y fechas correctamente.
    ''' </summary>
    ''' <param name="dbPath">Ruta al .sqlite</param>
    ''' <param name="tableName">Nombre de la tabla (opcional si fullQuery)</param>
    ''' <param name="whereClause">Cláusula WHERE sin la palabra WHERE (opcional)</param>
    ''' <param name="fullQuery">Consulta SQL completa (usa este si tableName vacío)</param>
    Public Shared Sub ExportToExcel(
            dbPath As String,
            Optional tableName As String = "",
            Optional whereClause As String = "",
            Optional fullQuery As String = "")

        ' 1) Construir SQL
        Dim sql As String
        If Not String.IsNullOrWhiteSpace(fullQuery) Then
            sql = fullQuery
        ElseIf Not String.IsNullOrWhiteSpace(tableName) Then
            sql = $"SELECT * FROM [{tableName}]"
            If Not String.IsNullOrWhiteSpace(whereClause) Then
                sql &= " WHERE " & whereClause
            End If
        Else
            Throw New ArgumentException("Debe indicar tableName/whereClause o fullQuery.")
        End If

        ' 2) Llenar DataTable
        Dim dt As New DataTable()
        Try
            Using conn As New SQLiteConnection($"Data Source={dbPath};Version=3;")
                conn.Open()
                Using cmd As New SQLiteCommand(sql, conn)
                    Using da As New SQLiteDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Forms.MessageBox.Show(
                $"Error al ejecutar consulta:{vbCrLf}{ex.Message}",
                "Error BD",
                Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Error)
            Return
        End Try

        If dt.Rows.Count = 0 Then
            Forms.MessageBox.Show(
                "La consulta no devolvió filas.",
                "Sin datos",
                Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Information)
            Return
        End If

        ' 3) Iniciar Excel
        Dim excelApp As Excel.Application = Nothing
        Dim wb As Excel.Workbook = Nothing
        Dim ws As Excel.Worksheet = Nothing

        Try
            excelApp = New Excel.Application With {
                .Visible = True,
                .DisplayAlerts = False
            }
            wb = excelApp.Workbooks.Add()
            ws = CType(wb.Sheets(1), Excel.Worksheet)
            ws.Name = If(String.IsNullOrWhiteSpace(tableName), "Query", tableName)

            ' 4) Determinar columnas a pintar (ocultar "id")
            Dim paintCols As New List(Of DataColumn)
            For Each dc As DataColumn In dt.Columns
                If Not String.Equals(dc.ColumnName, "id", StringComparison.OrdinalIgnoreCase) Then
                    paintCols.Add(dc)
                End If
            Next

            ' 5) Pintar encabezados
            For j As Integer = 0 To paintCols.Count - 1
                ws.Cells(1, j + 1).Value = paintCols(j).ColumnName
            Next
            Dim hdr As Excel.Range = ws.Range(ws.Cells(1, 1), ws.Cells(1, paintCols.Count))
            With hdr
                .Interior.Color = Drawing.ColorTranslator.ToOle(Drawing.Color.Navy)
                .Font.Color = Drawing.ColorTranslator.ToOle(Drawing.Color.White)
                .Font.Bold = True
                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            End With

            ' 6) Pintar filas y formatear según tipo
            For r As Integer = 0 To dt.Rows.Count - 1
                For j As Integer = 0 To paintCols.Count - 1
                    Dim dc = paintCols(j)
                    Dim cell = ws.Cells(r + 2, j + 1)
                    Dim val = dt.Rows(r)(dc)

                    cell.Value = val

                    ' Formato según tipo de columna
                    Select Case Type.GetTypeCode(dc.DataType)
                        Case TypeCode.Decimal, TypeCode.Double
                            ' numérico → moneda con 2 decimales
                            cell.NumberFormat = "$#,##0.00"
                            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                        Case TypeCode.Single, TypeCode.Int16, TypeCode.Int32, TypeCode.Int64
                            cell.NumberFormat = "@"
                            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                        Case TypeCode.DateTime
                            cell.NumberFormat = "dd\/mm\/yyyy"
                            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        Case Else
                            ' texto
                            cell.NumberFormat = "@"
                            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                    End Select
                Next
                ' fila alterna suave
                If (r Mod 2) = 1 Then
                    Dim rowR = ws.Range(ws.Cells(r + 2, 1), ws.Cells(r + 2, paintCols.Count))
                    rowR.Interior.Color = Drawing.ColorTranslator.ToOle(Drawing.Color.LightGray)
                End If
            Next

            ' 7) Ajustes finales
            ws.Columns.AutoFit()
            excelApp.ActiveWindow.SplitRow = 1
            excelApp.ActiveWindow.FreezePanes = True

        Catch ex As Exception
            Forms.MessageBox.Show(
                $"Error al exportar a Excel:{vbCrLf}{ex.Message}",
                "Error Excel",
                Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Error)

        Finally
            ' 8) Limpiar COM
            Try
                If ws IsNot Nothing Then Marshal.ReleaseComObject(ws)
                If wb IsNot Nothing Then Marshal.ReleaseComObject(wb)
                If excelApp IsNot Nothing Then Marshal.ReleaseComObject(excelApp)
            Catch : End Try

            ws = Nothing
            wb = Nothing
            excelApp = Nothing

            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

End Class
