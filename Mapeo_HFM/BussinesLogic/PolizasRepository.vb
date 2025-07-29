Imports System.Data
Imports System.Data.SQLite

Public Class PolizasRepository
    Private ReadOnly _connStr As String

    Public Sub New(rutaDB As String)
        _connStr = $"Data Source={rutaDB};Version=3;"
    End Sub

    Public Function GetEncabezado() As DataTable
        Using conn As New SQLiteConnection(_connStr)
            conn.Open()
            Dim cmd As New SQLiteCommand("SELECT * FROM polizas_encabezado LIMIT 1", conn)
            Dim dt As New DataTable()
            dt.Load(cmd.ExecuteReader())
            Return dt
        End Using
    End Function

    Public Function GetPolizasHFM() As DataTable
        Using conn As New SQLiteConnection(_connStr)
            conn.Open()
            Dim cmd As New SQLiteCommand("SELECT * FROM polizas_HFM", conn)
            Dim dt As New DataTable()
            dt.Load(cmd.ExecuteReader())
            Return dt
        End Using
    End Function

    Public Sub LimpiarTablas()
        Using conn As New SQLiteConnection(_connStr)
            conn.Open()
            Dim cmd1 As New SQLiteCommand("DELETE FROM polizas_encabezado", conn)
            cmd1.ExecuteNonQuery()
            Dim cmd2 As New SQLiteCommand("DELETE FROM polizas_HFM", conn)
            cmd2.ExecuteNonQuery()
            ' Resetear la secuencia de polizas_HFM
            Dim cmd3 As New SQLiteCommand("UPDATE sqlite_sequence SET seq = 0 WHERE name = 'polizas_HFM'", conn)
            cmd3.ExecuteNonQuery()
        End Using
    End Sub

    Public Sub InsertarEncabezado(row As DataRow)
        Using conn As New SQLiteConnection(_connStr)
            conn.Open()
            Dim cmd As New SQLiteCommand("INSERT INTO polizas_encabezado (Usuario, Fecha, Hora, Scenario, Year, Period, Value, currDate) VALUES (@Usuario, @Fecha, @Hora, @Scenario, @Year, @Period, @Value, @currDate)", conn)
            cmd.Parameters.AddWithValue("@Usuario", row("Usuario"))
            cmd.Parameters.AddWithValue("@Fecha", row("Fecha"))
            cmd.Parameters.AddWithValue("@Hora", row("Hora"))
            cmd.Parameters.AddWithValue("@Scenario", row("Scenario"))
            cmd.Parameters.AddWithValue("@Year", row("Year"))
            cmd.Parameters.AddWithValue("@Period", row("Period"))
            cmd.Parameters.AddWithValue("@Value", row("Value"))
            cmd.Parameters.AddWithValue("@currDate", row("currDate"))
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Public Sub InsertarPolizaHFM(row As DataRow)
        Using conn As New SQLiteConnection(_connStr)
            conn.Open()
            Dim cmd As New SQLiteCommand("INSERT INTO polizas_HFM (Grupo, Etiqueta, Descripcion, Entity, Account, Creado_por, Aprobado_por, Estado, Debe, Haber) VALUES (@Grupo, @Etiqueta, @Descripcion, @Entity, @Account, @CreadoPor, @AprobadoPor, @Estado, @Debe, @Haber)", conn)
            cmd.Parameters.AddWithValue("@Grupo", row("Grupo").ToString().Trim())
            cmd.Parameters.AddWithValue("@Etiqueta", row("Etiqueta").ToString().Trim())
            cmd.Parameters.AddWithValue("@Descripcion", row("Descripcion").ToString().Trim())
            cmd.Parameters.AddWithValue("@Entity", row("Entity").ToString().Trim())
            cmd.Parameters.AddWithValue("@Account", row("Account").ToString().Trim())
            cmd.Parameters.AddWithValue("@CreadoPor", row("Creado_por").ToString().Trim())
            cmd.Parameters.AddWithValue("@AprobadoPor", row("Aprobado_por").ToString().Trim())
            cmd.Parameters.AddWithValue("@Estado", row("Estado").ToString().Trim())
            cmd.Parameters.AddWithValue("@Debe", row("Debe"))
            cmd.Parameters.AddWithValue("@Haber", row("Haber"))
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Public Sub ActualizarPolizaHFM(row As DataRow)
        Using conn As New SQLiteConnection(_connStr)
            conn.Open()
            Dim cmd As New SQLiteCommand("UPDATE polizas_HFM SET Creado_por=@CreadoPor, Aprobado_por=@AprobadoPor, Estado=@Estado, Debe=@Debe, Haber=@Haber WHERE id=@id", conn)
            cmd.Parameters.AddWithValue("@CreadoPor", row("Creado_por"))
            cmd.Parameters.AddWithValue("@AprobadoPor", row("Aprobado_por"))
            cmd.Parameters.AddWithValue("@Estado", row("Estado"))
            cmd.Parameters.AddWithValue("@Debe", row("Debe"))
            cmd.Parameters.AddWithValue("@Haber", row("Haber"))
            cmd.Parameters.AddWithValue("@id", row("id"))
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Public Function BuscarPolizasSimilares(grupo As String, entity As String, account As String, excludeId As Integer) As DataTable
        Using conn As New SQLiteConnection(_connStr)
            conn.Open()
            Dim cmd As New SQLiteCommand("SELECT Debe, Haber FROM polizas_HFM WHERE Grupo=@Grupo AND Entity=@Entity AND Account=@Account AND id<>@id", conn)
            cmd.Parameters.AddWithValue("@Grupo", grupo)
            cmd.Parameters.AddWithValue("@Entity", entity)
            cmd.Parameters.AddWithValue("@Account", account)
            cmd.Parameters.AddWithValue("@id", excludeId)
            Dim dt As New DataTable()
            dt.Load(cmd.ExecuteReader())
            Return dt
        End Using
    End Function

    ' Métodos para respaldar, recuperar histórico, etc. pueden agregarse aquí.
    Public Function GetPolizasHFMFiltrado(Optional grupoFiltro As String = "", Optional entityFiltro As String = "", Optional descripFiltro As String = "") As DataTable
        Using conn As New SQLiteConnection(_connStr)
            conn.Open()
            Dim whereList As New List(Of String)
            Dim parametros As New List(Of SQLiteParameter)

            If Not String.IsNullOrWhiteSpace(grupoFiltro) Then
                whereList.Add("Grupo = @Grupo")
                parametros.Add(New SQLiteParameter("@Grupo", grupoFiltro))
            End If
            If Not String.IsNullOrWhiteSpace(entityFiltro) Then
                whereList.Add("Entity = @Entity")
                parametros.Add(New SQLiteParameter("@Entity", entityFiltro))
            End If
            If Not String.IsNullOrWhiteSpace(descripFiltro) Then
                whereList.Add("Descripcion = @Descripcion")
                parametros.Add(New SQLiteParameter("@Descripcion", descripFiltro))
            End If

            Dim whereClause As String = ""
            If whereList.Count > 0 Then
                whereClause = " WHERE " & String.Join(" AND ", whereList)
            End If

            Dim sql As String = "SELECT *, '' AS Semaforo FROM polizas_HFM" & whereClause

            Dim dt As New DataTable()
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddRange(parametros.ToArray())
                Using reader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
            Return dt
        End Using
    End Function
End Class