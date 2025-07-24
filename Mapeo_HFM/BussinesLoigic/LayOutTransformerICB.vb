

Imports System.Data.SQLite

    ''' <summary>
    ''' Clase responsable de extraer datos de la tabla pre_lay_out,
    ''' transformarlos según reglas de negocio y cargarlos en lay_out.
    ''' </summary>
    Public Class LayOutTransformerICB
        Private ReadOnly _connectionString As String

    ''' <summary>
    ''' Constructor que recibe la ruta del archivo SQLite.
    ''' </summary>
    ''' <param name="dbFilePath">Ruta al archivo de base de datos SQLite.</param>
    Public Sub New(dbFilePath As String)
        ' Construye la cadena de conexión a partir de la ruta del archivo
        _connectionString = $"Data Source={dbFilePath};Version=3;"
    End Sub
    ''' <summary>
    ''' Ejecuta la transformación de pre_lay_out a lay_out.
    ''' </summary>
    Public Sub Transform()
            Using conn As New SQLiteConnection(_connectionString)
                conn.Open()

                ' 1) Obtener el período máximo para el campo MES
                Dim maxPeriodo As Integer
            Using cmdMax As New SQLiteCommand("SELECT COALESCE(MAX(periodo),0) FROM pre_lay_out", conn)
                Dim result = cmdMax.ExecuteScalar()
                maxPeriodo = If(result IsNot Nothing AndAlso Not Convert.IsDBNull(result), Convert.ToInt32(result), 0)
            End Using

            ' 1.1) Limpiar tabla lay_out y resetear secuencia de ID autoincremental
            Using cmdDelete As New SQLiteCommand("DELETE FROM lay_out;", conn)
                cmdDelete.ExecuteNonQuery()
            End Using
            Using cmdResetSeq As New SQLiteCommand("DELETE FROM sqlite_sequence WHERE name='lay_out';", conn)
                cmdResetSeq.ExecuteNonQuery()
            End Using

            ' 2) Iniciar transacción para operaciones de escritura
            Using transaction = conn.BeginTransaction()

                ' 3) Leer todos los registros de pre_lay_out
                Dim sqlSelect = "SELECT grupo, cuenta_sap, cuenta_oracle, descripcion_cuenta_sific, ICIA_SIFIC, ejercicio, periodo, saldo_acum FROM pre_lay_out"
                Using cmdSelect As New SQLiteCommand(sqlSelect, conn, transaction)
                        Using reader As SQLiteDataReader = cmdSelect.ExecuteReader()
                            While reader.Read()
                                Dim cia As String = reader("grupo").ToString()
                                Dim s_neg As Integer = 2000
                                Dim cta As String = reader("cuenta_oracle").ToString()
                                Dim descrip As String = reader("descripcion_cuenta_sific").ToString()
                            Dim icp As String = reader("ICIA_SIFIC").ToString()
                            Dim mon As String = "MX"
                                Dim tipo As String = "H"
                                Dim yearVal As Integer = Convert.ToInt32(reader("ejercicio"))
                                Dim mes As Integer = maxPeriodo
                                Dim importe As Decimal = Convert.ToDecimal(reader("saldo_acum"))

                                ' 4) Calcular TOP según catálogo cat_TOP
                                Dim topVal As Integer = 0
                                If Not String.IsNullOrWhiteSpace(icp) Then
                                    Dim sqlTop = "SELECT TOP FROM cat_TOP " &
                                             "WHERE cuentaSAP = @CuentaSAP " &
                                             "AND ICP = @ICP AND grupo = @Grupo"
                                    Using cmdTop As New SQLiteCommand(sqlTop, conn, transaction)
                                        cmdTop.Parameters.AddWithValue("@CuentaSAP", reader("cuenta_sap").ToString())
                                        cmdTop.Parameters.AddWithValue("@ICP", icp)
                                        cmdTop.Parameters.AddWithValue("@Grupo", cia)
                                        Dim objTop = cmdTop.ExecuteScalar()
                                        If objTop IsNot Nothing AndAlso Not Convert.IsDBNull(objTop) Then
                                            topVal = Convert.ToInt32(objTop)
                                        End If
                                    End Using
                                End If

                                ' 5) Insertar en lay_out
                                Dim sqlInsert = "INSERT INTO lay_out (CIA, S_NEG, CTA, DESCRIP, ICP, MON, TIPO, YEAR, MES, TOP, IMPORTE) " &
                                            "VALUES (@CIA, @S_NEG, @CTA, @DESCRIP, @ICP, @MON, @TIPO, @YEAR, @MES, @TOP, @IMPORTE)"
                                Using cmdInsert As New SQLiteCommand(sqlInsert, conn, transaction)
                                    cmdInsert.Parameters.AddWithValue("@CIA", cia)
                                    cmdInsert.Parameters.AddWithValue("@S_NEG", s_neg)
                                    cmdInsert.Parameters.AddWithValue("@CTA", cta)
                                    cmdInsert.Parameters.AddWithValue("@DESCRIP", descrip)
                                    cmdInsert.Parameters.AddWithValue("@ICP", icp)
                                    cmdInsert.Parameters.AddWithValue("@MON", mon)
                                    cmdInsert.Parameters.AddWithValue("@TIPO", tipo)
                                    cmdInsert.Parameters.AddWithValue("@YEAR", yearVal)
                                    cmdInsert.Parameters.AddWithValue("@MES", mes)
                                    cmdInsert.Parameters.AddWithValue("@TOP", topVal)
                                    cmdInsert.Parameters.AddWithValue("@IMPORTE", importe)
                                    cmdInsert.ExecuteNonQuery()
                                End Using
                            End While
                        End Using
                    End Using

                    ' 6) Confirmar transacción
                    transaction.Commit()
                End Using

                conn.Close()
            End Using
        End Sub
    End Class
