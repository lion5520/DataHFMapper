Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SQLite
Imports System.Linq

Public Class AperturaDetalleProcessor

    Private ReadOnly _rutaBd As String

    ''' <summary>
    ''' Inicializa la clase con la ruta de la BD SQLite.
    ''' </summary>
    Public Sub New(rutaBd As String)
        If String.IsNullOrWhiteSpace(rutaBd) Then
            Throw New ArgumentException("La ruta de la BD no puede estar vacía.", NameOf(rutaBd))
        End If
        If Not File.Exists(rutaBd) Then
            Throw New FileNotFoundException("No se encontró el archivo de base de datos.", rutaBd)
        End If
        _rutaBd = rutaBd
    End Sub

    ''' <summary>
    ''' Ejecuta todo el proceso de apertura de detalle de saldos.
    ''' </summary>
    Public Sub ProcesarReporteIC()
        Using conn As New SQLiteConnection($"Data Source={_rutaBd}")
            conn.Open()
            Using tran = conn.BeginTransaction()

                ' --------------------------------------------------------
                ' 1) Traer los registros “crudos” de reporte_IC (incluye Saldo)
                '    filtrando solo los pares que sí tienen padre en t_in_sap
                ' --------------------------------------------------------
                Dim dtRep As New DataTable()
                Dim sqlRep As String = "
SELECT
  LTRIM(r.ICSap,       '0') AS ICSap,
  LTRIM(r.SociedadSap, '0') AS SociedadSap,
  LTRIM(r.CuentaSap,   '0') AS CuentaSap,
  r.Saldo                     AS Saldo,
  r.Cuenta_Parte_Relacionada  AS CtaOracle
FROM reporte_IC AS r
JOIN t_in_sap AS t
  ON LTRIM(t.sociedad, '0') = LTRIM(r.SociedadSap, '0')
 AND t.numero_cuenta      = LTRIM(r.CuentaSap,   '0');
"
                Using cmdRep As New SQLiteCommand(sqlRep, conn, tran)
                    Using da As New SQLiteDataAdapter(cmdRep)
                        da.Fill(dtRep)
                    End Using
                End Using

                ' DEBUG: Asegúrate de que la columna "Saldo" existe
                Debug.WriteLine("Columnas en dtRep: " &
                    String.Join(", ", dtRep.Columns.Cast(Of DataColumn)().Select(Function(c) c.ColumnName)))

                ' --------------------------------------------------------
                ' 2) Prepara las columnas de t_in_sap para insertar copias
                ' --------------------------------------------------------
                Dim cols = GetTableColumns(conn, tran, "t_in_sap") _
                    .Where(Function(c) c <> "rowid" AndAlso c <> "id") _
                    .ToList()

                For Each reg As DataRow In dtRep.Rows
                    Dim ic As String = reg.Field(Of String)("ICSap")
                    Dim soc As String = reg.Field(Of String)("SociedadSap")
                    Dim cta As String = reg.Field(Of String)("CuentaSap")
                    Dim saldo As Double = reg.Field(Of Double)("Saldo")

                    Dim dtPadre As New DataTable()
                    Dim sqlPadre As String = "SELECT rowid AS RowId, * FROM t_in_sap WHERE LTRIM(sociedad,'0')=@soc AND LTRIM(numero_cuenta,'0')=@cta LIMIT 1;"
                    Using cmdPadre As New SQLiteCommand(sqlPadre, conn, tran)
                        cmdPadre.Parameters.AddWithValue("@soc", soc)
                        cmdPadre.Parameters.AddWithValue("@cta", cta)
                        Using daPadre As New SQLiteDataAdapter(cmdPadre)
                            daPadre.Fill(dtPadre)
                        End Using
                    End Using
                    If dtPadre.Rows.Count = 0 Then Continue For

                    Dim padre = dtPadre.Rows(0)

                    Dim colNames = String.Join(", ", cols)
                    Dim paramNames = String.Join(", ", cols.Select(Function(c) "@" & c))
                    Dim sqlIns = $"INSERT INTO t_in_sap ({colNames}) VALUES ({paramNames});"
                    Using cmdIns As New SQLiteCommand(sqlIns, conn, tran)
                        For Each col In cols
                            cmdIns.Parameters.AddWithValue("@" & col, padre(col))
                        Next
                        cmdIns.Parameters("@" & "deudor_acreedor_2").Value = ic
                        cmdIns.Parameters("@" & "saldo_acum").Value = saldo
                        cmdIns.ExecuteNonQuery()
                    End Using

                    ' ----------------------------------------------
                    ' Paso 2: Reclasificación
                    ' ----------------------------------------------
                    Dim sqlClas As String = "SELECT Entidad_i, CUENTA_i, Tipo_i FROM cat_deudor_acredor WHERE ICP_i=@ic AND CUENTA_d=@cta LIMIT 1;"
                    Dim dtClas As New DataTable()
                    Using cmdClas As New SQLiteCommand(sqlClas, conn, tran)
                        cmdClas.Parameters.AddWithValue("@ic", ic)
                        cmdClas.Parameters.AddWithValue("@cta", cta)
                        Using daClas As New SQLiteDataAdapter(cmdClas)
                            daClas.Fill(dtClas)
                        End Using
                    End Using
                    If dtClas.Rows.Count > 0 Then
                        Dim clas = dtClas.Rows(0)
                        Dim destSoc As String = LTrim(clas("Entidad_i").ToString())
                        Dim destCta As String = clas("CUENTA_i").ToString()
                        Dim tipo As String = clas("Tipo_i").ToString().Trim().ToUpperInvariant()

                        Dim sqlBuscaDest As String = "SELECT rowid AS RowId, saldo_acum FROM t_in_sap WHERE LTRIM(sociedad,'0')=@soc AND (deudor_acreedor_2='[ICP None]' OR deudor_acreedor_2=@ic) LIMIT 1;"
                        Dim dtDest As New DataTable()
                        Using cmdDest As New SQLiteCommand(sqlBuscaDest, conn, tran)
                            cmdDest.Parameters.AddWithValue("@soc", destSoc)
                            cmdDest.Parameters.AddWithValue("@ic", ic)
                            Using daDest As New SQLiteDataAdapter(cmdDest)
                                daDest.Fill(dtDest)
                            End Using
                        End Using

                        If dtDest.Rows.Count > 0 Then
                            Dim dest = dtDest.Rows(0)
                            Dim saldoOrig As Double = Convert.ToDouble(dest("saldo_acum"))
                            Dim nuevoSaldo As Double = If(tipo = "C", saldoOrig - saldo, saldoOrig + saldo)
                            Using cmdUpd As New SQLiteCommand("UPDATE t_in_sap SET saldo_acum=@s, numero_cuenta=@cta WHERE rowid=@rid;", conn, tran)
                                cmdUpd.Parameters.AddWithValue("@s", nuevoSaldo)
                                cmdUpd.Parameters.AddWithValue("@cta", destCta)
                                cmdUpd.Parameters.AddWithValue("@rid", dest.Field(Of Long)("RowId"))
                                cmdUpd.ExecuteNonQuery()
                            End Using
                        End If

                        ' --------------------------------------
                        ' Paso 3: Bitácora en polizas_HFM
                        ' --------------------------------------
                        Dim grupo As String = String.Empty
                        Using cmdGrp As New SQLiteCommand("SELECT GRUPO FROM GL_ICP_Grupos WHERE GL_ICP=@key LIMIT 1;", conn, tran)
                            cmdGrp.Parameters.AddWithValue("@key", destCta)
                            Dim res = cmdGrp.ExecuteScalar()
                            If res IsNot Nothing Then grupo = res.ToString()
                        End Using

                        Using cmdBit As New SQLiteCommand("INSERT INTO polizas_HFM (Grupo, Descripcion, Account, Debe, Haber) VALUES (@grp,'RECLACIFICACION',@acc,@deb,@hab);", conn, tran)
                            cmdBit.Parameters.AddWithValue("@grp", grupo)
                            cmdBit.Parameters.AddWithValue("@acc", destCta)
                            If tipo = "C" Then
                                cmdBit.Parameters.AddWithValue("@deb", saldo)
                                cmdBit.Parameters.AddWithValue("@hab", DBNull.Value)
                            Else
                                cmdBit.Parameters.AddWithValue("@deb", DBNull.Value)
                                cmdBit.Parameters.AddWithValue("@hab", saldo)
                            End If
                            cmdBit.ExecuteNonQuery()
                        End Using
                    End If
                Next

                tran.Commit()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Recupera la lista de nombres de columnas de una tabla SQLite.
    ''' </summary>
    Private Function GetTableColumns(
        conn As SQLiteConnection,
        tran As SQLiteTransaction,
        tableName As String) _
    As List(Of String)

        Dim lista As New List(Of String)()
        Using cmd As New SQLiteCommand($"PRAGMA table_info({tableName});", conn, tran)
            Using rdr = cmd.ExecuteReader()
                While rdr.Read()
                    lista.Add(rdr.GetString(rdr.GetOrdinal("name")))
                End While
            End Using
        End Using
        Return lista
    End Function


End Class
