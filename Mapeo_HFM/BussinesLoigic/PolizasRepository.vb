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
            Dim cmd As New SQLiteCommand("select * from polizas_HFM where not grupo is null and grupo <> '' and not Entity  is null and Entity <> '' ", conn)
            Dim dt As New DataTable()
            dt.Load(cmd.ExecuteReader())
            Return dt
        End Using
    End Function

    Public Sub LimpiarTablas()
        Using conn As New SQLiteConnection(_connStr)
            conn.Open()
            Using tran As SQLiteTransaction = conn.BeginTransaction()
                Try
                    Dim cmd1 As New SQLiteCommand("DELETE FROM polizas_encabezado", conn, tran)
                    cmd1.ExecuteNonQuery()
                    Dim cmd2 As New SQLiteCommand("DELETE FROM polizas_HFM", conn, tran)
                    cmd2.ExecuteNonQuery()
                    Dim cmd3 As New SQLiteCommand("UPDATE sqlite_sequence SET seq = 0 WHERE name = 'polizas_HFM'", conn, tran)
                    cmd3.ExecuteNonQuery()
                    tran.Commit()
                Catch ex As Exception
                    tran.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Sub

    Public Sub InsertarEncabezado(row As DataRow)
        Using conn As New SQLiteConnection(_connStr)
            conn.Open()
            Using tran As SQLiteTransaction = conn.BeginTransaction()
                Try
                    Dim cmd As New SQLiteCommand("INSERT INTO polizas_encabezado (Usuario, Fecha, Hora, Scenario, Year, Period, Value, currDate) VALUES (@Usuario, @Fecha, @Hora, @Scenario, @Year, @Period, @Value, @currDate)", conn, tran)
                    cmd.Parameters.AddWithValue("@Usuario", row("Usuario"))
                    cmd.Parameters.AddWithValue("@Fecha", row("Fecha"))
                    cmd.Parameters.AddWithValue("@Hora", row("Hora"))
                    cmd.Parameters.AddWithValue("@Scenario", row("Scenario"))
                    cmd.Parameters.AddWithValue("@Year", row("Year"))
                    cmd.Parameters.AddWithValue("@Period", row("Period"))
                    cmd.Parameters.AddWithValue("@Value", row("Value"))
                    cmd.Parameters.AddWithValue("@currDate", row("currDate"))
                    cmd.ExecuteNonQuery()
                    tran.Commit()
                Catch ex As Exception
                    tran.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Sub

    Public Sub InsertarPolizaHFM(row As DataRow)
        Using conn As New SQLiteConnection(_connStr)
            conn.Open()
            Using tran As SQLiteTransaction = conn.BeginTransaction()
                Try
                    Dim cmd As New SQLiteCommand("INSERT INTO polizas_HFM (Grupo, Etiqueta, Descripcion, Entity, Account, Creado_por, Aprobado_por, Estado, Debe, Haber) VALUES (@Grupo, @Etiqueta, @Descripcion, @Entity, @Account, @CreadoPor, @AprobadoPor, @Estado, @Debe, @Haber)", conn, tran)
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
                    tran.Commit()
                Catch ex As Exception
                    tran.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Sub

    Public Sub ActualizarPolizaHFM(row As DataRow)
        Using conn As New SQLiteConnection(_connStr)
            conn.Open()
            Using tran As SQLiteTransaction = conn.BeginTransaction()
                Try
                    Dim cmd As New SQLiteCommand("UPDATE polizas_HFM SET Creado_por=@CreadoPor, Aprobado_por=@AprobadoPor, Estado=@Estado, Debe=@Debe, Haber=@Haber WHERE id=@id", conn, tran)
                    cmd.Parameters.AddWithValue("@CreadoPor", row("Creado_por"))
                    cmd.Parameters.AddWithValue("@AprobadoPor", row("Aprobado_por"))
                    cmd.Parameters.AddWithValue("@Estado", row("Estado"))
                    cmd.Parameters.AddWithValue("@Debe", row("Debe"))
                    cmd.Parameters.AddWithValue("@Haber", row("Haber"))
                    cmd.Parameters.AddWithValue("@id", row("id"))
                    cmd.ExecuteNonQuery()
                    tran.Commit()
                Catch ex As Exception
                    tran.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Sub

    Public Sub ActualizarSaldoPolizaHFM(idSeleccionado As Integer, saldo As Decimal)
        Using conn As New SQLiteConnection(_connStr)
            conn.Open()
            Using tran As SQLiteTransaction = conn.BeginTransaction()
                Try
                    Dim cmd As New SQLiteCommand("UPDATE polizas_HFM SET saldo = @Saldo WHERE id = @Id", conn, tran)
                    cmd.Parameters.AddWithValue("@Saldo", saldo)
                    cmd.Parameters.AddWithValue("@Id", idSeleccionado)
                    cmd.ExecuteNonQuery()
                    tran.Commit()
                Catch ex As Exception
                    tran.Rollback()
                    Throw
                End Try
            End Using
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

    ' Leer todos los registros de t_in_sap
    Public Function GetTInSap() As DataTable
        Using conn As New SQLiteConnection(_connStr)
            conn.Open()
            Dim cmd As New SQLiteCommand("SELECT * FROM t_in_sap where asignacion = ''  ", conn)
            Dim dt As New DataTable()
            dt.Load(cmd.ExecuteReader())
            Return dt
        End Using
    End Function

    ' Buscar en polizas_HFM por sociedad y numero_cuenta
    Public Function BuscarPolizaPorSociedadCuenta(sociedad As String, cuenta As String) As DataRow
        Using conn As New SQLiteConnection(_connStr)
            conn.Open()
            Dim cmd As New SQLiteCommand("SELECT * FROM polizas_HFM WHERE Entity = @Sociedad AND Account = @Cuenta", conn)
            cmd.Parameters.AddWithValue("@Sociedad", sociedad)
            cmd.Parameters.AddWithValue("@Cuenta", cuenta)
            Using reader = cmd.ExecuteReader()
                If reader.Read() Then
                    Dim dt As New DataTable()
                    dt.Load(reader)
                    If dt.Rows.Count > 0 Then Return dt.Rows(0)
                End If
            End Using
        End Using
        Return Nothing
    End Function

    ' Buscar registros por grupo y descripcion
    Public Function BuscarPolizasPorGrupoDescripcion(grupo As String, descripcion As String) As DataTable
        Using conn As New SQLiteConnection(_connStr)
            conn.Open()
            Dim cmd As New SQLiteCommand("SELECT * FROM polizas_HFM WHERE Grupo = @Grupo AND Descripcion = @Descripcion ORDER BY id ASC", conn)
            cmd.Parameters.AddWithValue("@Grupo", grupo)
            cmd.Parameters.AddWithValue("@Descripcion", descripcion)
            Dim dt As New DataTable()
            dt.Load(cmd.ExecuteReader())
            Return dt
        End Using
    End Function

    ' Actualizar registro en t_in_sap
    Public Sub ActualizarTInSapAsignacion(id As Integer, asignacion As String, Optional saldo_acum As Decimal? = Nothing)
        Using conn As New SQLiteConnection(_connStr)
            conn.Open()
            Using tran As SQLiteTransaction = conn.BeginTransaction()
                Try
                    Dim sql As String = "UPDATE t_in_sap SET asignacion = @Asignacion"
                    If saldo_acum.HasValue Then
                        sql &= ", saldo_acum = @SaldoAcum"
                    End If
                    sql &= " WHERE id = @Id"
                    Dim cmd As New SQLiteCommand(sql, conn, tran)
                    cmd.Parameters.AddWithValue("@Asignacion", asignacion)
                    If saldo_acum.HasValue Then
                        cmd.Parameters.AddWithValue("@SaldoAcum", saldo_acum.Value)
                    End If
                    cmd.Parameters.AddWithValue("@Id", id)
                    cmd.ExecuteNonQuery()
                    tran.Commit()
                Catch ex As Exception
                    tran.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Sub

    ' Insertar nuevo registro en t_in_sap
    Public Sub InsertarTInSap(sociedad As String, numero_cuenta As String, cuenta_mayor_hfm As String, saldo_acum As Decimal, periodo As String, ejercicio As String, asignacion As String)
        Using conn As New SQLiteConnection(_connStr)
            conn.Open()
            Using tran As SQLiteTransaction = conn.BeginTransaction()
                Try
                    Dim cmd As New SQLiteCommand("INSERT INTO t_in_sap (sociedad, numero_cuenta, cuenta_mayor_hfm, saldo_acum, periodo, ejercicio, asignacion) VALUES (@Sociedad, @NumeroCuenta, @CuentaMayorHFM, @SaldoAcum, @Periodo, @Ejercicio, @Asignacion)", conn, tran)
                    cmd.Parameters.AddWithValue("@Sociedad", sociedad)
                    cmd.Parameters.AddWithValue("@NumeroCuenta", numero_cuenta)
                    cmd.Parameters.AddWithValue("@CuentaMayorHFM", cuenta_mayor_hfm)
                    cmd.Parameters.AddWithValue("@SaldoAcum", saldo_acum)
                    cmd.Parameters.AddWithValue("@Periodo", periodo)
                    cmd.Parameters.AddWithValue("@Ejercicio", ejercicio)
                    cmd.Parameters.AddWithValue("@Asignacion", asignacion)
                    cmd.ExecuteNonQuery()
                    tran.Commit()
                Catch ex As Exception
                    tran.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Sub

    Public Sub ClonarTablaTInSapOriginal()
        Using conn As New SQLiteConnection(_connStr)
            conn.Open()
            Using tran As SQLiteTransaction = conn.BeginTransaction()
                Try
                    ' 1. Validar si existe la tabla t_in_sap_original
                    Dim existeTabla As Boolean = False
                    Using cmdCheck As New SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table' AND name='t_in_sap_original';", conn, tran)
                        Using reader = cmdCheck.ExecuteReader()
                            existeTabla = reader.HasRows
                        End Using
                    End Using

                    ' 1.1 Si existe, realizar DROP
                    If existeTabla Then
                        Using cmdDrop As New SQLiteCommand("DROP TABLE t_in_sap_original;", conn, tran)
                            cmdDrop.ExecuteNonQuery()
                        End Using
                    End If

                    ' 1.2 Clonar la tabla t_in_sap
                    Using cmdClone As New SQLiteCommand("CREATE TABLE t_in_sap_original AS SELECT * FROM t_in_sap;", conn, tran)
                        cmdClone.ExecuteNonQuery()
                    End Using

                    tran.Commit()
                Catch ex As Exception
                    tran.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Sub

    Public Function GetSumariaValidacionData() As DataTable
        Dim dt As New DataTable()
        '    Dim sql As String = "
        '    select t.agrupador_detalle, t.cuenta_oracle, t.descripcion_cuenta_oracle,  t.numero_cuenta, t.texto_explicativo, 
        '       sum(t.saldo_acum) as suma_cias , sum(t.saldo_acum - s.saldo_acum) as reclasificacion,  0 as eliminacion,  sum(s.saldo_acum) as saldo_final
        '    from t_in_sap_original t, t_in_sap s
        '    where s.id = t.id
        '    group by t.agrupador_detalle, t.cuenta_oracle, t.descripcion_cuenta_oracle,  t.numero_cuenta, t.texto_explicativo
        '    order by t.agrupador_detalle, t.cuenta_oracle, t.descripcion_cuenta_oracle,  t.numero_cuenta, t.texto_explicativo;

        Dim sql As String = "
select COALESCE(t.agrupador_detalle,'') AS agrupador_detalle, COALESCE(t.cuenta_oracle,'') AS cuenta_oracle, 
   COALESCE(t.descripcion_cuenta_oracle,'') AS descripcion_cuenta_oracle,  t.numero_cuenta,  COALESCE(t.texto_explicativo,'') AS texto_explicativo, 
   sum(COALESCE( s.saldo_acum,0)) as suma_cias , sum(COALESCE( t.saldo_acum,0) - COALESCE( s.saldo_acum,0)) as reclasificacion,  0 as eliminacion,  
   sum(COALESCE( t.saldo_acum,0)) as saldo_final
from t_in_sap t
   LEFT join t_in_sap_original s on t.id = s.id
group by t.agrupador_detalle, t.cuenta_oracle, t.descripcion_cuenta_oracle,  t.numero_cuenta, t.texto_explicativo
order by t.agrupador_detalle, t.cuenta_oracle, t.descripcion_cuenta_oracle,  t.numero_cuenta, t.texto_explicativo; 
        "
        Using conn As New SQLite.SQLiteConnection(_connStr)
            conn.Open()
            Using cmd As New SQLite.SQLiteCommand(Sql, conn)
                Using da As New SQLite.SQLiteDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function
End Class