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

                'Obtiene las operaciones complementarias 
                Dim dtRep As New DataTable()
                Using cmdRep As New SQLiteCommand(
                    "SELECT LTRIM(ICSap,'0') AS ICSap, " &
                    "       LTRIM(SociedadSap,'0') AS SociedadSap, " &
                    "       CuentaSap, " &
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

                'Por cada registro, busca a su registro padre [ICP None]
                For Each repRow As DataRow In dtRep.Rows
                    Dim soc As String = NormalizeKey(repRow("SociedadSap").ToString())
                    Dim cta As String = NormalizeKey(repRow("CuentaSap").ToString())
                    Dim ic As String = NormalizeKey(repRow("ICSap").ToString())
                    Dim saldo As Double = Convert.ToDouble(repRow("Saldo"))
                    Dim ctaOracle As String = repRow("Cuenta_Parte_Relacionada").ToString()

                    Dim dtPadre As New DataTable()
                    Using cmdFind As New SQLiteCommand(
                        "SELECT rowid AS RowId, * FROM t_in_sap " &
                        "WHERE LTRIM(sociedad,'0')=@soc AND numero_cuenta=@cta AND deudor_acreedor_2='[ICP None]';", conn, tran)
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


                    'Busca cuenta equivalente de Reclasificación de la operacion 
                    Dim dtClas As New DataTable()
                    Dim sqlClas As String

                    sqlClas = "SELECT LTRIM(Entidad_i,'0'), CUENTA_i, ICP_i, Tipo_i AS Operacion_Destino FROM cat_deudor_acredor " &
                                  "WHERE LTRIM(Entidad_d,'0')=@soc AND LTRIM(ICP_i,'0')=@ic AND CUENTA_d=@cta ;"

                    Using cmdClas As New SQLiteCommand(sqlClas, conn, tran)
                        cmdClas.Parameters.AddWithValue("@soc", soc)
                        cmdClas.Parameters.AddWithValue("@ic", ic)
                        cmdClas.Parameters.AddWithValue("@cta", cta)
                        Using da As New SQLiteDataAdapter(cmdClas)
                            da.Fill(dtClas)
                        End Using
                    End Using


                    Dim Valida As Boolean = True
                    Dim ctaDest As String = ""
                    Dim operacionDest As String = ""
                    If dtClas.Rows.Count > 0 Then
                        Dim clas_cuentaEquivalente = dtClas.Rows(0)
                        ctaDest = NormalizeKey(clas_cuentaEquivalente("CUENTA_i").ToString())
                        operacionDest = NormalizeKey(clas_cuentaEquivalente("Operacion_Destino").ToString())
                    Else
                        ctaDest = "Reclasific N/D"
                        Valida = False
                    End If


                    'Obtiene datos de cuenta de reclacificacion si existen en lay out orinial 
                    ' Variables para almacenar los datos retorno
                    Dim textoExplicativo As String = String.Empty
                    Dim cuentaMayorHfm As String = String.Empty
                    Dim descripcionCuentaSific As String = String.Empty
                    Dim descripcionCuentaOracle As String = String.Empty
                    Dim agrupadorTipo As String = String.Empty
                    Dim agrupadorCuenta As String = String.Empty

                    ' Preparamos el comando para traer solo los campos deseados
                    Dim dtCampos As New DataTable()
                    Using cmdFind As New SQLiteCommand(
                                        "SELECT 
                                            texto_explicativo, 
                                            cuenta_mayor_hfm, 
                                            descripcion_cuenta_sific, 
                                            descripcion_cuenta_oracle, 
                                            agrupador_tipo, 
                                            agrupador_cuenta 
                                         FROM t_in_sap 
                                         WHERE numero_cuenta = @cta;", conn, tran)
                        cmdFind.Parameters.AddWithValue("@cta", ctaDest)

                        Using da As New SQLiteDataAdapter(cmdFind)
                            da.Fill(dtCampos)
                        End Using
                    End Using

                    ' Si existe al menos un registro, asignamos cada campo a su variable
                    If dtCampos.Rows.Count > 0 Then
                        Dim row = dtCampos.Rows(0)
                        textoExplicativo = row.Field(Of String)("texto_explicativo")
                        cuentaMayorHfm = row.Field(Of String)("cuenta_mayor_hfm")
                        descripcionCuentaSific = row.Field(Of String)("descripcion_cuenta_sific")
                        descripcionCuentaOracle = row.Field(Of String)("descripcion_cuenta_oracle")
                        agrupadorTipo = row.Field(Of String)("agrupador_tipo")
                        agrupadorCuenta = row.Field(Of String)("agrupador_cuenta")
                    Else
                        Valida = False
                    End If

                    ' ---- ---------------   ----------------  ---


                    ' Insertar copia ajustando IC y saldo
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
                        cmdIns.Parameters("@numero_cuenta").Value = ctaDest
                        cmdIns.Parameters("@texto_explicativo").Value = $"Reclacific N/A Desc {cta}-{ctaDest}"
                        cmdIns.Parameters("@asignacion").Value = $"Reclacificacion {cta}-{ctaDest}"

                        If Valida Then
                            cmdIns.Parameters("@texto_explicativo").Value = textoExplicativo
                            cmdIns.Parameters("@cuenta_mayor_hfm").Value = cuentaMayorHfm
                            cmdIns.Parameters("@descripcion_cuenta_sific").Value = descripcionCuentaSific
                            cmdIns.Parameters("@descripcion_cuenta_oracle").Value = descripcionCuentaOracle
                            cmdIns.Parameters("@agrupador_tipo").Value = agrupadorTipo
                            cmdIns.Parameters("@agrupador_cuenta").Value = agrupadorCuenta
                        End If


                        cmdIns.ExecuteNonQuery()
                        newRowId = conn.LastInsertRowId
                    End Using

                    'Actualiza saldo en registro padre, quitando la poarte proporcional de la oepracion integrada
                    Dim nuevoSaldoPadre As Double = padre("saldo_acum") - saldo

                    Using cmdUpdPadre As New SQLiteCommand(
                            "UPDATE t_in_sap SET saldo_acum=@s WHERE rowid=@rid;", conn, tran)
                        cmdUpdPadre.Parameters.AddWithValue("@s", nuevoSaldoPadre)
                        cmdUpdPadre.Parameters.AddWithValue("@rid", padre("RowId"))
                        cmdUpdPadre.ExecuteNonQuery()
                    End Using


                    'Agrega operacion en Bitácora en polizas_HFM
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
                        If String.Equals(operacionDest, "C", StringComparison.OrdinalIgnoreCase) Then
                            cmdBit.Parameters.AddWithValue("@deb", saldo)
                            cmdBit.Parameters.AddWithValue("@hab", DBNull.Value)
                        Else
                            cmdBit.Parameters.AddWithValue("@deb", DBNull.Value)
                                cmdBit.Parameters.AddWithValue("@hab", saldo)
                            End If
                            cmdBit.ExecuteNonQuery()
                        End Using

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
