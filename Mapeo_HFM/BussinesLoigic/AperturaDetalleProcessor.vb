Imports System
Imports System.Data
Imports System.Data.SQLite
Imports System.IO
Imports System.Linq

Public Class AperturaDetalleProcessor

    Private ReadOnly _rutaBd As String

    ''' <summary>
    ''' Elimina ceros a la izquierda en claves de cuenta o sociedad.
    ''' </summary>
    Private Shared Function NormalizeKey(value As String) As String
        Return value?.TrimStart("0"c)
    End Function

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
    ''' Ejecuta el proceso completo solicitado de apertura y reclasificación.
    ''' </summary>
    Public Sub ProcesarReporteIC()
        Using conn As New SQLiteConnection($"Data Source={_rutaBd}")
            conn.Open()
            Using tran = conn.BeginTransaction()

                Dim dtRep As New DataTable()
                Using cmdRep As New SQLiteCommand(
                    "SELECT LTRIM(ICSap,'0') AS ICSap, " &
                    "       LTRIM(SociedadSap,'0') AS SociedadSap, " &
                    "       LTRIM(CuentaSap,'0') AS CuentaSap, " &
                    "       Cuenta_Parte_Relacionada, " &
                    "       Saldo " &
                    "  FROM reporte_IC", conn, tran)
                    Using da As New SQLiteDataAdapter(cmdRep)
                        da.Fill(dtRep)
                    End Using
                End Using

                Dim cols = GetTableColumns(conn, tran, "t_in_sap") _
                               .Where(Function(c) c <> "rowid" AndAlso c <> "id") _
                               .ToList()

                ' Verificar si cat_deudor_acredor cuenta con la columna Operacion_Destino
                Dim hasOperacion As Boolean = ColumnExists(conn, tran, "cat_deudor_acredor", "Operacion_Destino")

                For Each repRow As DataRow In dtRep.Rows
                    Dim soc As String = NormalizeKey(repRow("SociedadSap").ToString())
                    Dim cta As String = NormalizeKey(repRow("CuentaSap").ToString())
                    Dim ic As String = NormalizeKey(repRow("ICSap").ToString())
                    Dim saldo As Double = Convert.ToDouble(repRow("Saldo"))
                    Dim ctaOracle As String = repRow("Cuenta_Parte_Relacionada").ToString()

                    Dim dtPadre As New DataTable()
                    Using cmdFind As New SQLiteCommand(
                        "SELECT rowid AS RowId, * FROM t_in_sap " &
                        "WHERE LTRIM(sociedad,'0')=@soc AND LTRIM(numero_cuenta,'0')=@cta LIMIT 1;", conn, tran)
                        cmdFind.Parameters.AddWithValue("@soc", soc)
                        cmdFind.Parameters.AddWithValue("@cta", cta)
                        Using da As New SQLiteDataAdapter(cmdFind)
                            da.Fill(dtPadre)
                        End Using
                    End Using
                    If dtPadre.Rows.Count = 0 Then
                        Continue For
                    End If
                    Dim padre = dtPadre.Rows(0)

                    ' Paso 1: Insertar copia ajustando IC y saldo
                    Dim colNames = String.Join(", ", cols)
                    Dim paramNames = String.Join(", ", cols.Select(Function(c) "@" & c))
                    Dim sqlIns = $"INSERT INTO t_in_sap ({colNames}) VALUES ({paramNames});"
                    Dim newRowId As Long
                    Using cmdIns As New SQLiteCommand(sqlIns, conn, tran)
                        For Each col In cols
                            cmdIns.Parameters.AddWithValue("@" & col, padre(col))
                        Next
                        cmdIns.Parameters("@deudor_acreedor_2").Value = ic
                        cmdIns.Parameters("@saldo_acum").Value = saldo
                        cmdIns.Parameters("@cuenta_oracle").Value = ctaOracle
                        cmdIns.ExecuteNonQuery()
                        newRowId = conn.LastInsertRowId
                    End Using

                    ' Paso 2: Reclasificación
                    Dim dtClas As New DataTable()
                    Dim sqlClas As String
                    If hasOperacion Then
                        sqlClas = "SELECT Entidad_i, CUENTA_i, Operacion_Destino FROM cat_deudor_acredor " &
                                  "WHERE LTRIM(ICP_i,'0')=@ic AND LTRIM(CUENTA_d,'0')=@cta LIMIT 1;"
                    Else
                        sqlClas = "SELECT Entidad_i, CUENTA_i, Tipo_i AS Operacion_Destino FROM cat_deudor_acredor " &
                                  "WHERE LTRIM(ICP_i,'0')=@ic AND LTRIM(CUENTA_d,'0')=@cta LIMIT 1;"
                    End If
                    Using cmdClas As New SQLiteCommand(sqlClas, conn, tran)
                        cmdClas.Parameters.AddWithValue("@ic", ic)
                        cmdClas.Parameters.AddWithValue("@cta", cta)
                        Using da As New SQLiteDataAdapter(cmdClas)
                            da.Fill(dtClas)
                        End Using
                    End Using

                    If dtClas.Rows.Count > 0 Then
                        Dim clas = dtClas.Rows(0)
                        Dim socDest As String = NormalizeKey(clas("Entidad_i").ToString())
                        Dim ctaDest As String = NormalizeKey(clas("CUENTA_i").ToString())
                        Dim operacion As String = clas("Operacion_Destino").ToString()

                        Dim dtDest As New DataTable()
                        Using cmdDest As New SQLiteCommand(
                            "SELECT rowid AS RowId, saldo_acum FROM t_in_sap " &
                            "WHERE LTRIM(sociedad,'0')=@sd AND (deudor_acreedor_2=@ic OR deudor_acreedor_2='ICP_NONE') " &
                            "LIMIT 1;", conn, tran)
                            cmdDest.Parameters.AddWithValue("@sd", socDest)
                            cmdDest.Parameters.AddWithValue("@ic", ic)
                            Using da As New SQLiteDataAdapter(cmdDest)
                                da.Fill(dtDest)
                            End Using
                        End Using

                        If dtDest.Rows.Count > 0 Then
                            Dim destRow = dtDest.Rows(0)
                            Dim saldoActual As Double = Convert.ToDouble(destRow("saldo_acum"))
                            Dim nuevoSaldo As Double = saldoActual
                            If String.Equals(operacion, "C", StringComparison.OrdinalIgnoreCase) Then
                                nuevoSaldo -= saldo
                            Else
                                nuevoSaldo += saldo
                            End If

                            Using cmdUpd As New SQLiteCommand(
                                "UPDATE t_in_sap SET saldo_acum=@s WHERE rowid=@rid;", conn, tran)
                                cmdUpd.Parameters.AddWithValue("@s", nuevoSaldo)
                                cmdUpd.Parameters.AddWithValue("@rid", destRow("RowId"))
                                cmdUpd.ExecuteNonQuery()
                            End Using

                            Using cmdUpdNew As New SQLiteCommand(
                                "UPDATE t_in_sap SET numero_cuenta=@cta, sociedad=@soc WHERE rowid=@rid;", conn, tran)
                                cmdUpdNew.Parameters.AddWithValue("@cta", ctaDest)
                                cmdUpdNew.Parameters.AddWithValue("@soc", socDest)
                                cmdUpdNew.Parameters.AddWithValue("@rid", newRowId)
                                cmdUpdNew.ExecuteNonQuery()
                            End Using
                        End If

                        ' Ajustar saldo del registro padre
                        Dim saldoPadre As Double = Convert.ToDouble(padre("saldo_acum"))
                        Dim nuevoSaldoPadre As Double = saldoPadre
                        If String.Equals(operacion, "C", StringComparison.OrdinalIgnoreCase) Then
                            nuevoSaldoPadre -= saldo
                        Else
                            nuevoSaldoPadre += saldo
                        End If

                        Using cmdUpdPadre As New SQLiteCommand(
                            "UPDATE t_in_sap SET saldo_acum=@s WHERE rowid=@rid;", conn, tran)
                            cmdUpdPadre.Parameters.AddWithValue("@s", nuevoSaldoPadre)
                            cmdUpdPadre.Parameters.AddWithValue("@rid", padre("RowId"))
                            cmdUpdPadre.ExecuteNonQuery()
                        End Using

                        ' Paso 3: Bitácora en polizas_HFM
                        Dim grupo As String = String.Empty
                        Using cmdGrupo As New SQLiteCommand("SELECT GRUPO FROM GL_ICP_Grupos WHERE GL_ICP=@key LIMIT 1;", conn, tran)
                            cmdGrupo.Parameters.AddWithValue("@key", ctaDest)
                            Dim val = cmdGrupo.ExecuteScalar()
                            If val IsNot Nothing Then grupo = val.ToString()
                        End Using

                        Using cmdBit As New SQLiteCommand(
                            "INSERT INTO polizas_HFM (Grupo, Descripcion, Account, Debe, Haber) " &
                            "VALUES (@grp,'RECLACIFICACION',@acc,@deb,@hab);", conn, tran)
                            cmdBit.Parameters.AddWithValue("@grp", grupo)
                            cmdBit.Parameters.AddWithValue("@acc", ctaDest)
                            If String.Equals(operacion, "C", StringComparison.OrdinalIgnoreCase) Then
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

    Private Function GetTableColumns(conn As SQLiteConnection,
                                     tran As SQLiteTransaction,
                                     tableName As String) As List(Of String)
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

    ''' <summary>
    ''' Verifica si una tabla contiene determinada columna.
    ''' </summary>
    Private Function ColumnExists(conn As SQLiteConnection,
                                   tran As SQLiteTransaction,
                                   tableName As String,
                                   columnName As String) As Boolean
        Using cmd As New SQLiteCommand($"PRAGMA table_info({tableName});", conn, tran)
            Using rdr = cmd.ExecuteReader()
                While rdr.Read()
                    Dim name = rdr.GetString(rdr.GetOrdinal("name"))
                    If String.Equals(name, columnName, StringComparison.OrdinalIgnoreCase) Then
                        Return True
                    End If
                End While
            End Using
        End Using
        Return False
    End Function

End Class
