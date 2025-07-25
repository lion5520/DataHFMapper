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
                ' 2) Agrupar en código por SociedadSap + CuentaSap
                ' --------------------------------------------------------
                Dim grupos = dtRep.AsEnumerable().
                             GroupBy(Function(r) New With {
                                 Key .Soc = r.Field(Of String)("SociedadSap"),
                                 Key .Cta = r.Field(Of String)("CuentaSap")
                             })

                For Each grupo In grupos
                    Dim soc = grupo.Key.Soc
                    Dim cta = grupo.Key.Cta
                    ' Aquí sí existe la columna "Saldo"
                    Dim detalles = grupo.ToList()

                    ' Verificar pares con saldos iguales
                    Dim filasAInsertar As New List(Of DataRow)()
                    If detalles.Count = 2 Then
                        Dim s1 = Math.Round(detalles(0).Field(Of Double)("Saldo"), 2)
                        Dim s2 = Math.Round(detalles(1).Field(Of Double)("Saldo"), 2)
                        If s1 <> s2 Then
                            filasAInsertar.AddRange(detalles)
                        End If
                    Else
                        filasAInsertar.AddRange(detalles)
                    End If

                    Dim totalSaldo = filasAInsertar.Sum(Function(r) r.Field(Of Double)("Saldo"))

                    ' ----------------------------------------------------
                    ' 3) Buscar registro padre en t_in_sap
                    ' ----------------------------------------------------
                    Dim dtPadre As New DataTable()
                    Dim sqlPadre As String = "
SELECT rowid AS RowId, *
FROM t_in_sap
WHERE LTRIM(sociedad,'0') = @soc
  AND numero_cuenta      = @cta
  AND deudor_acreedor_2  = '[ICP None]';
"
                    Using cmdPadre As New SQLiteCommand(sqlPadre, conn, tran)
                        cmdPadre.Parameters.AddWithValue("@soc", soc)
                        cmdPadre.Parameters.AddWithValue("@cta", cta)
                        Using daPadre As New SQLiteDataAdapter(cmdPadre)
                            daPadre.Fill(dtPadre)
                        End Using
                    End Using
                    Dim restarPadre As Boolean = True
                    If dtPadre.Rows.Count = 0 Then
                        ' Si no hay registro padre con [ICP None], tomamos cualquiera
                        restarPadre = False
                        Dim sqlAny As String = "SELECT rowid AS RowId, * FROM t_in_sap WHERE LTRIM(sociedad,'0') = @soc AND numero_cuenta = @cta LIMIT 1;"
                        Using cmdAny As New SQLiteCommand(sqlAny, conn, tran)
                            cmdAny.Parameters.AddWithValue("@soc", soc)
                            cmdAny.Parameters.AddWithValue("@cta", cta)
                            Using daAny As New SQLiteDataAdapter(cmdAny)
                                daAny.Fill(dtPadre)
                            End Using
                        End Using
                    End If

                    If dtPadre.Rows.Count = 0 Then
                        ' Sin filas de referencia, no podemos insertar
                        Continue For
                    End If

                    Dim padre = dtPadre.Rows(0)
                    Dim padreRowId As Long = padre.Field(Of Long)("RowId")

                    ' ----------------------------------------------------
                    ' 4) Preparar inserción dinámica en t_in_sap
                    ' ----------------------------------------------------
                    ' Obtenemos todas las columnas de la tabla t_in_sap
                    Dim cols = GetTableColumns(conn, tran, "t_in_sap") _
               .Where(Function(c) c <> "rowid" AndAlso c <> "id") _
               .ToList()
                    ' Ahora cols no contiene ni rowid ni id

                    ' Insertar o actualizar un detalle por cada fila localizada
                    For Each detalle In filasAInsertar
                        Dim ic As String = detalle.Field(Of String)("ICSap")
                        Dim saldoDet As Double = Math.Round(detalle.Field(Of Double)("Saldo"), 2)
                        Dim ctaOra As String = detalle.Field(Of String)("CtaOracle")

                        If String.IsNullOrWhiteSpace(ctaOra) Then
                            ' Se asume reclasificación, se ignora
                            Continue For
                        End If

                        Dim dtExist As New DataTable()
                        Using cmdE As New SQLiteCommand("SELECT rowid AS RowId FROM t_in_sap WHERE LTRIM(sociedad,'0')=@soc AND numero_cuenta=@cta AND deudor_acreedor_2=@ic;", conn, tran)
                            cmdE.Parameters.AddWithValue("@soc", soc)
                            cmdE.Parameters.AddWithValue("@cta", cta)
                            cmdE.Parameters.AddWithValue("@ic", ic)
                            Using daE As New SQLiteDataAdapter(cmdE)
                                daE.Fill(dtExist)
                            End Using
                        End Using

                        If dtExist.Rows.Count > 0 Then
                            ' Actualizar saldo y cuenta_oracle
                            Dim ridExist As Long = dtExist.Rows(0).Field(Of Long)("RowId")
                            Using cmdUpdDet As New SQLiteCommand("UPDATE t_in_sap SET saldo_acum=@s, cuenta_oracle=@co WHERE rowid=@rid;", conn, tran)
                                cmdUpdDet.Parameters.AddWithValue("@s", saldoDet)
                                cmdUpdDet.Parameters.AddWithValue("@co", ctaOra)
                                cmdUpdDet.Parameters.AddWithValue("@rid", ridExist)
                                cmdUpdDet.ExecuteNonQuery()
                            End Using
                        Else
                            Dim colNames = String.Join(", ", cols)
                            Dim paramNames = String.Join(", ", cols.Select(Function(c) "@" & c))
                            Dim sqlIns = $"INSERT INTO t_in_sap ({colNames}) VALUES ({paramNames});"

                            Using cmdIns As New SQLiteCommand(sqlIns, conn, tran)
                                For Each col In cols
                                    cmdIns.Parameters.AddWithValue("@" & col, padre(col))
                                Next

                                cmdIns.Parameters("@" & "deudor_acreedor_2").Value = ic
                                cmdIns.Parameters("@" & "saldo_acum").Value = saldoDet
                                cmdIns.Parameters("@" & "cuenta_oracle").Value = ctaOra

                                cmdIns.ExecuteNonQuery()
                            End Using
                        End If
                    Next


                    ' ----------------------------------------------------
                    ' 5) Ajustar el registro padre restándole el total
                    ' ----------------------------------------------------
                    If restarPadre Then
                        Dim saldoOriginal = Convert.ToDouble(padre("saldo_acum"))
                        Dim nuevoSaldo = Math.Round(saldoOriginal - totalSaldo, 2)
                        Using cmdUpd As New SQLiteCommand("
UPDATE t_in_sap
   SET saldo_acum = @ns
 WHERE rowid      = @rid;", conn, tran)
                            cmdUpd.Parameters.AddWithValue("@ns", nuevoSaldo)
                            cmdUpd.Parameters.AddWithValue("@rid", padreRowId)
                            cmdUpd.ExecuteNonQuery()
                        End Using
                    End If

                Next ' siguiente grupo

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
