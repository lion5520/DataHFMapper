Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SQLite
Imports System.Windows.Forms
Imports System.Linq

Public Class UPD_matriz_costo_ingreso
    Private ReadOnly _dbPath As String
    Private ReadOnly _excelPath As String

    ''' <param name="dbPath">Ruta al archivo .sqlite</param>
    ''' <param name="excelPath">Ruta al archivo Excel a importar</param>
    Public Sub New(dbPath As String, excelPath As String)
        If String.IsNullOrWhiteSpace(dbPath) Then Throw New ArgumentException("Ruta BD vacía", NameOf(dbPath))
        If Not File.Exists(dbPath) Then Throw New FileNotFoundException("No existe BD", dbPath)
        If String.IsNullOrWhiteSpace(excelPath) Then Throw New ArgumentException("Ruta Excel vacía", NameOf(excelPath))
        If Not File.Exists(excelPath) Then Throw New FileNotFoundException("No existe archivo Excel", excelPath)
        _dbPath = dbPath
        _excelPath = excelPath
    End Sub

    ''' <summary>
    ''' Importa la pestaña BASE del Excel a la tabla costo_ingreso_acum.
    ''' </summary>
    Public Sub Importar()
        ' 1) Cargar toda la hoja BASE en un DataTable (HDR=NO)
        Dim dtExcel As New DataTable()
        Const sheetName As String = "BASE$"
        Dim connStr = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={_excelPath};Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=1;'"
        Try
            Using oledb As New System.Data.OleDb.OleDbConnection(connStr)
                oledb.Open()
                Dim schema = oledb.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, Nothing)
                If Not schema.AsEnumerable().
                       Any(Function(r) String.Equals(r("TABLE_NAME").ToString().Trim(), sheetName, StringComparison.OrdinalIgnoreCase)) Then
                    MessageBox.Show("La hoja 'BASE' no existe.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
                Using cmd As New System.Data.OleDb.OleDbCommand($"SELECT * FROM [{sheetName}]", oledb)
                    Using da As New System.Data.OleDb.OleDbDataAdapter(cmd)
                        da.Fill(dtExcel)
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error leyendo Excel: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try

        ' 2) Definir mapeo encabezado Excel -> campo DB
        Dim mapping As New Dictionary(Of String, String) From {
            {"SOC. SAP", "SOC_SAP"}, {"DEUDOR ACREEDOR", "tipo_deu_acre"}, {"No. ORACLE", "ICP_ORACLE"},
            {"Sociedad SAP", "ICP_SAP"}, {"INTERCOMPAÑIA", "nombre_ICP"}, {"GRUPO IC", "grupo_de_ICP"},
            {"AÑO", "year"}, {"MES", "mes"}, {"MONEDA", "moneda"},
            {"Monto Total Moneda Origen", "monto_t_mon_origen"}, {"Monto Total MXP", "MONTO_T_MXP"},
            {"DESCRIPCION DE SERVICIO", "servicio_descrip"}, {"CUENTA SAP", "CUENTA_SAP"},
            {"DESCRIPCION CUENTA SAP", "descrip_cuenta_sap"}, {"CUENTA ORACLE", "CUENTA_ORACLE"},
            {"DESCRIPCION CUENTA ORACLE", "descrip_cuenta_oracle"},
            {"Clasificacion COSTO Ó INGRESO", "clasifc_cost_ingreso"}, {"TOP", "TOP"}
        }

        ' 3) Encontrar fila de encabezado buscando "SOC. SAP" en columna C
        Dim headerRowIndex As Integer = -1
        Dim colIndexC = ColumnLetterToNumber("C")
        For i As Integer = 0 To dtExcel.Rows.Count - 1
            Dim val = Convert.ToString(dtExcel.Rows(i)(colIndexC))
            If String.Equals(val?.Trim(), "SOC. SAP", StringComparison.OrdinalIgnoreCase) Then
                headerRowIndex = i
                Exit For
            End If
        Next
        If headerRowIndex < 0 Then
            MessageBox.Show("No se encontró la fila de encabezados (SOC. SAP).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' 4) Determinar posición de cada encabezado en la fila headerRowIndex
        Dim positions As New Dictionary(Of String, Integer)
        Dim hdrRow = dtExcel.Rows(headerRowIndex)
        For Each hdrText In mapping.Keys
            For c As Integer = 0 To dtExcel.Columns.Count - 1
                If String.Equals(Convert.ToString(hdrRow(c))?.Trim(), hdrText, StringComparison.OrdinalIgnoreCase) Then
                    positions(hdrText) = c
                    Exit For
                End If
            Next
        Next

        ' 5) Insertar en SQLite
        Using conn As New SQLiteConnection($"Data Source={_dbPath}")
            conn.Open()
            ' Limpiar tabla y resetear autoincrement
            Using clearCmd = New SQLiteCommand(
                    "DELETE FROM costo_ingreso_acum;" &
                    "DELETE FROM sqlite_sequence WHERE name='costo_ingreso_acum';",
                    conn)
                clearCmd.ExecuteNonQuery()
            End Using

            Dim dbCols = mapping.Values.ToList()
            Dim paramList = String.Join(",", dbCols.Select(Function(c) "@" & c))
            Dim sql = $"INSERT INTO costo_ingreso_acum ({String.Join(",", dbCols)}) VALUES ({paramList});"
            Using cmd As New SQLiteCommand(sql, conn)
                For Each c In dbCols
                    cmd.Parameters.Add(New SQLiteParameter("@" & c, DbType.String))
                Next

                For r As Integer = headerRowIndex + 1 To dtExcel.Rows.Count - 1
                    Dim row = dtExcel.Rows(r)
                    ' Omitir filas vacías
                    If mapping.Keys.All(Function(h) String.IsNullOrWhiteSpace(Convert.ToString(row(positions(h))))) Then
                        Continue For
                    End If
                    ' Asignar parámetros
                    For Each kv In mapping
                        Dim raw = row(positions(kv.Key))
                        cmd.Parameters("@" & kv.Value).Value = If(raw Is Nothing OrElse raw Is DBNull.Value, DBNull.Value, raw)
                    Next
                    cmd.ExecuteNonQuery()
                Next
            End Using
        End Using

        MessageBox.Show("Importación completada.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ''' <summary>
    ''' Convierte letra de columna Excel (A,B,...,Z,AA,...) a índice 0-based.
    ''' </summary>
    Private Function ColumnLetterToNumber(col As String) As Integer
        Dim sum As Integer = 0
        For Each ch As Char In col
            sum = sum * 26 + (AscW(ch) - AscW("A"c) + 1)
        Next
        Return sum - 1
    End Function
End Class
