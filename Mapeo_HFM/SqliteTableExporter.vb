Imports System
Imports System.IO
Imports System.Data.SQLite
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Excel = Microsoft.Office.Interop.Excel

Public Class SqliteTableExporter

    Private ReadOnly _dbPath As String
    Private ReadOnly _tableName As String
    Private ReadOnly _defaultOut As String

    ''' <summary>
    ''' dbPath: ruta al archivo .sqlite  
    ''' tableName: nombre de la tabla a exportar  
    ''' defaultOutputFile: ruta sugerida (incluye nombre.xlsx) para guardar
    ''' </summary>
    Public Sub New(dbPath As String, tableName As String, defaultOutputFile As String)
        If String.IsNullOrWhiteSpace(dbPath) Then Throw New ArgumentException("BD vacía", NameOf(dbPath))
        If Not File.Exists(dbPath) Then Throw New FileNotFoundException("No existe BD", dbPath)
        If String.IsNullOrWhiteSpace(tableName) Then Throw New ArgumentException("Tabla vacía", NameOf(tableName))
        If String.IsNullOrWhiteSpace(defaultOutputFile) Then Throw New ArgumentException("Ruta sugerida vacía", NameOf(defaultOutputFile))
        _dbPath = dbPath
        _tableName = tableName
        _defaultOut = defaultOutputFile
    End Sub

    ''' <summary>
    ''' Muestra un SaveFileDialog para que el usuario elija dónde guardar,
    ''' y luego exporta la tabla a .xlsx con formato.
    ''' </summary>
    Public Sub Export()
        ' 0) Pide ruta al usuario
        Dim outputFile As String = Nothing
        Using dlg As New SaveFileDialog()
            dlg.Title = $"Guardar '{_tableName}' como Excel"
            dlg.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
            dlg.FileName = Path.GetFileName(_defaultOut)
            dlg.InitialDirectory = If(Path.GetDirectoryName(_defaultOut), Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments))
            If dlg.ShowDialog() <> DialogResult.OK Then
                Return  ' usuario canceló
            End If
            outputFile = dlg.FileName
        End Using

        Dim dt As New System.Data.DataTable()

        ' 1) Leer esquema y datos
        Using conn As New SQLiteConnection($"Data Source={_dbPath}")
            conn.Open()

            ' 1.a) PRAGMA para obtener columnas (excluye PK)
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
                MessageBox.Show($"La tabla '{_tableName}' no tiene columnas exportables.",
                                "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' 1.b) SELECT datos
            Dim sql = $"SELECT {String.Join(", ", colNames)} FROM {_tableName};"
            Using da As New SQLiteDataAdapter(sql, conn)
                da.Fill(dt)
            End Using
        End Using

        ' 2) Preparar array para volcar en bloque
        Dim rows = dt.Rows.Count
        Dim cols = dt.Columns.Count
        Dim arr(rows, cols - 1) As Object

        ' 2.a) Encabezados
        For c As Integer = 0 To cols - 1
            arr(0, c) = dt.Columns(c).ColumnName
        Next
        ' 2.b) Datos
        For r As Integer = 0 To rows - 1
            For c As Integer = 0 To cols - 1
                Dim v = dt.Rows(r)(c)
                arr(r + 1, c) = If(v Is DBNull.Value, "", v)
            Next
        Next

        ' 3) Volcar a Excel via COM Interop (una sola llamada)
        Dim xlApp As Excel.Application = Nothing
        Dim xlWb As Excel.Workbook = Nothing
        Dim xlWs As Excel.Worksheet = Nothing

        Try
            xlApp = New Excel.Application With {.Visible = False, .DisplayAlerts = False}
            xlWb = xlApp.Workbooks.Add()
            xlWs = CType(xlWb.Sheets(1), Excel.Worksheet)
            xlWs.Name = _tableName

            Dim startCell = xlWs.Cells(1, 1)
            Dim endCell = xlWs.Cells(rows + 1, cols)
            Dim writeRange = xlWs.Range(startCell, endCell)
            writeRange.Value2 = arr

            ' 4) Formato
            ' 4.a) Negrita en fila de encabezados
            Dim hdr = xlWs.Range(xlWs.Cells(1, 1), xlWs.Cells(1, cols))
            hdr.Font.Bold = True

            ' 4.b) Autoajustar columnas
            For c As Integer = 1 To cols
                CType(xlWs.Columns(c), Excel.Range).AutoFit()
            Next

            ' 5) Guardar
            xlWb.SaveAs(outputFile)
            xlWb.Close()
            xlApp.Quit()

            MessageBox.Show($"Tabla '{_tableName}' exportada con éxito a:{vbCrLf}{outputFile}",
                            "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
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
