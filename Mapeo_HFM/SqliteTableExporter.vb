Imports System
Imports System.IO
Imports System.Data.SQLite
Imports System.Runtime.InteropServices
Imports Excel = Microsoft.Office.Interop.Excel

Public Class SqliteTableExporter

    Private ReadOnly _dbPath As String
    Private ReadOnly _tableName As String
    Private ReadOnly _outputFile As String

    ''' <summary>
    ''' dbPath: ruta al archivo .sqlite  
    ''' tableName: nombre de la tabla a exportar  
    ''' outputFile: ruta completa donde guardar el .xlsx
    ''' </summary>
    Public Sub New(dbPath As String, tableName As String, outputFile As String)
        If String.IsNullOrWhiteSpace(dbPath) Then Throw New ArgumentException("BD vacía", NameOf(dbPath))
        If Not File.Exists(dbPath) Then Throw New FileNotFoundException("No existe BD", dbPath)
        If String.IsNullOrWhiteSpace(tableName) Then Throw New ArgumentException("Tabla vacía", NameOf(tableName))
        If String.IsNullOrWhiteSpace(outputFile) Then Throw New ArgumentException("Archivo salida vacío", NameOf(outputFile))
        _dbPath = dbPath
        _tableName = tableName
        _outputFile = outputFile
    End Sub

    Public Sub Export()
        Dim dt As New System.Data.DataTable()

        ' 1) Leer esquema y datos
        Using conn As New SQLiteConnection($"Data Source={_dbPath}")
            conn.Open()

            ' 1.a) PRAGMA para columnas
            Dim colNames As New System.Collections.Generic.List(Of String)
            Using cmdInfo As New SQLiteCommand($"PRAGMA table_info({_tableName});", conn)
                Using r = cmdInfo.ExecuteReader()
                    While r.Read()
                        If Convert.ToInt32(r("pk")) = 0 Then
                            colNames.Add(r("name").ToString())
                        End If
                    End While
                End Using
            End Using

            If colNames.Count = 0 Then
                Throw New InvalidOperationException($"'{_tableName}' no tiene columnas exportables.")
            End If

            ' 1.b) SELECT datos
            Dim sql = $"SELECT {String.Join(", ", colNames)} FROM {_tableName};"
            Using da As New SQLiteDataAdapter(sql, conn)
                da.Fill(dt)
            End Using
        End Using

        ' 2) Preparar array (filas + 1 cabecera) × columnas
        Dim rows = dt.Rows.Count
        Dim cols = dt.Columns.Count
        Dim arr(,) As Object = New Object(rows, cols - 1) {}

        ' 2.a) Cabeceras
        For c As Integer = 0 To cols - 1
            arr(0, c) = dt.Columns(c).ColumnName
        Next
        ' 2.b) Datos
        For r As Integer = 0 To rows - 1
            For c As Integer = 0 To cols - 1
                Dim val = dt.Rows(r)(c)
                arr(r + 1, c) = If(val Is Nothing OrElse val Is DBNull.Value, "", val)
            Next
        Next

        ' 3) Volcar a Excel via COM Interop con un único Value2
        Dim xlApp As Excel.Application = Nothing
        Dim xlWb As Excel.Workbook = Nothing
        Dim xlWs As Excel.Worksheet = Nothing

        Try
            xlApp = New Excel.Application With {.Visible = False, .DisplayAlerts = False}
            xlWb = xlApp.Workbooks.Add
            xlWs = CType(xlWb.Sheets(1), Excel.Worksheet)
            xlWs.Name = _tableName

            ' Rango destino: A1 hasta (rows+1, cols)
            Dim endCell = xlWs.Cells(rows + 1, cols)
            Dim startCell = xlWs.Cells(1, 1)
            Dim writeRange = xlWs.Range(startCell, endCell)
            writeRange.Value2 = arr  ' volcado masivo

            ' 4) Formateo
            ' 4.a) Negrita en cabecera
            Dim hdrRange = xlWs.Range(xlWs.Cells(1, 1), xlWs.Cells(1, cols))
            hdrRange.Font.Bold = True

            ' 4.b) Autoajustar columnas
            For c As Integer = 1 To cols
                CType(xlWs.Columns(c), Excel.Range).AutoFit()
            Next

            ' 5) Guardar y cerrar
            xlWb.SaveAs(_outputFile)
            xlWb.Close()
            xlApp.Quit()

            MessageBox.Show($"Tabla '{_tableName}' exportada a:{vbCrLf}{_outputFile}",
                            "Exportación completada", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Finally
            ' 6) Liberar COM
            If xlWs IsNot Nothing Then Marshal.ReleaseComObject(xlWs)
            If xlWb IsNot Nothing Then Marshal.ReleaseComObject(xlWb)
            If xlApp IsNot Nothing Then Marshal.ReleaseComObject(xlApp)
            xlWs = Nothing
            xlWb = Nothing
            xlApp = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

End Class
