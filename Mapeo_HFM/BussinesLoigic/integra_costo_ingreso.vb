Imports System
Imports System.Data
Imports System.Data.SQLite
Imports System.IO
Imports System.Linq
Imports System.Windows.Forms

Public Class integra_costo_ingreso

    Private ReadOnly _dbPath As String

    Public Sub New(dbPath As String)
        If String.IsNullOrWhiteSpace(dbPath) Then
            Throw New ArgumentException("La ruta de la base de datos no puede estar vacía.", NameOf(dbPath))
        End If
        If Not File.Exists(dbPath) Then
            Throw New FileNotFoundException("No se encontró el archivo de base de datos.", dbPath)
        End If
        _dbPath = dbPath
    End Sub

    ''' <summary>
    ''' Ejecuta la rutina de integración de costo_ingreso_acum en t_in_sap.
    ''' </summary>
    Public Sub Procesar()
        Using conn As New SQLiteConnection($"Data Source={_dbPath}")
            conn.Open()
            Using tran = conn.BeginTransaction()
                ' Paso 1: obtener lista de grupos
                Dim grupos As New List(Of String)
                Dim sqlGr As String =
                    "SELECT DISTINCT g.GRUPO " &
                    "FROM costo_ingreso_acum ci " &
                    "  JOIN GL_ICP_Grupos g ON LTRIM(ci.SOC_SAP,'0')=LTRIM(g.GL_ICP,'0') " &
                    "ORDER BY g.GRUPO;"
                Using cmdGr As New SQLiteCommand(sqlGr, conn, tran)
                    Using rdr = cmdGr.ExecuteReader()
                        While rdr.Read()
                            grupos.Add(rdr.GetString(0))
                        End While
                    End Using
                End Using

                If grupos.Count = 0 Then
                    MessageBox.Show("No se encontraron grupos para procesar.", "Atención",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                ' Paso 2: seleccionar grupo
                Dim prompt = "Grupos disponibles: " & String.Join(", ", grupos) & vbCrLf &
                             "Ingrese el grupo a procesar:"
                Dim grupoSel = InputBox(prompt, "Seleccione Grupo").Trim()
                If String.IsNullOrWhiteSpace(grupoSel) OrElse Not grupos.Contains(grupoSel) Then
                    MessageBox.Show("Grupo inválido o cancelado.", "Atención",
                                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Return
                End If

                ' Paso 3: obtener sociedades para el grupo
                Dim socList As New List(Of String)
                Dim sqlSoc As String =
                    "SELECT DISTINCT LTRIM(ci.SOC_SAP,'0') " &
                    "FROM costo_ingreso_acum ci " &
                    "  JOIN GL_ICP_Grupos g ON LTRIM(ci.SOC_SAP,'0')=LTRIM(g.GL_ICP,'0') " &
                    "WHERE g.GRUPO=@grp;"
                Using cmdSoc As New SQLiteCommand(sqlSoc, conn, tran)
                    cmdSoc.Parameters.AddWithValue("@grp", grupoSel)
                    Using rdr = cmdSoc.ExecuteReader()
                        While rdr.Read()
                            socList.Add(rdr.GetString(0))
                        End While
                    End Using
                End Using

                If socList.Count = 0 Then
                    MessageBox.Show("No hay sociedades asociadas al grupo.", "Atención",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                ' Paso 4: cargar filas de costo_ingreso_acum para esas sociedades
                Dim dtCI As New DataTable()
                Dim inClause = String.Join(",", socList.Select(Function(s) $"'{s}'"))
                Dim sqlCI = $"SELECT * FROM costo_ingreso_acum WHERE LTRIM(SOC_SAP,'0') IN ({inClause});"
                Using daCI As New SQLiteDataAdapter(sqlCI, conn)
                    daCI.Fill(dtCI)
                End Using

                ' Paso 5: procesar cada fila leyendo y actualizando ICP None
                While dtCI.Rows.Count > 0
                    Dim rowCI = dtCI.Rows(0)
                    Dim soc = TrimLeadingZeros(Convert.ToString(rowCI("SOC_SAP")))
                    Dim cta = TrimLeadingZeros(Convert.ToString(rowCI("CUENTA_SAP")))
                    Dim icp = TrimLeadingZeros(Convert.ToString(rowCI("ICP_ORACLE")))
                    Dim monto = Math.Round(Convert.ToDouble(rowCI("MONTO_T_MXP")), 2, MidpointRounding.AwayFromZero)


                    ' 5.1) Leer o crear registro base “[ICP None]”
                    Dim dtBase As New DataTable()
                    Using cmdBase As New SQLiteCommand(
                        "SELECT ID, saldo_acum FROM t_in_sap " &
                        "WHERE LTRIM(sociedad,'0')=@soc " &
                          "AND LTRIM(numero_cuenta,'0')=@cta " &
                          "AND deudor_acreedor_2='[ICP None]';",
                        conn, tran)
                        cmdBase.Parameters.AddWithValue("@soc", soc)
                        cmdBase.Parameters.AddWithValue("@cta", cta)
                        Using daBase As New SQLiteDataAdapter(cmdBase)
                            daBase.Fill(dtBase)
                        End Using
                    End Using

                    'Si no hay registro IPC_NOONE no hay donde acumular diferencia i es el caso se crea
                    If dtBase.Rows.Count = 0 Then
                        Using cmdNewBase As New SQLiteCommand(
                            "INSERT INTO t_in_sap " &
                            "(sociedad, numero_cuenta, deudor_acreedor_2, saldo_acum, descripcion_cuenta_sific, asignacion) " &
                            "VALUES (@soc,@cta,'[ICP None]',0,'Inicial ICP None',@desc);",
                            conn, tran)
                            cmdNewBase.Parameters.AddWithValue("@soc", soc)
                            cmdNewBase.Parameters.AddWithValue("@cta", cta)
                            cmdNewBase.Parameters.AddWithValue("@desc", "Arraste de saldo diferencial para la cuenta  " & soc & "-" & cta)
                            cmdNewBase.ExecuteNonQuery()
                        End Using
                        dtBase.Clear()
                        Using cmdBase2 As New SQLiteCommand(
                            "SELECT ID, saldo_acum FROM t_in_sap WHERE rowid = last_insert_rowid();",
                            conn, tran)
                            Using da2 As New SQLiteDataAdapter(cmdBase2)
                                da2.Fill(dtBase)
                            End Using
                        End Using
                    End If

                    Dim baseId = dtBase.Rows(0).Field(Of Long)("ID")
                    Dim baseSaldo = Math.Round(Convert.ToDouble(dtBase.Rows(0)("saldo_acum")), 2, MidpointRounding.AwayFromZero)


                    ' 5.2) Insertar renglón de operación con ICP_ORACLE
                    Using cmdIns As New SQLiteCommand(
                        "INSERT INTO t_in_sap " &
                        "(sociedad, numero_cuenta, deudor_acreedor_2, saldo_acum, asignacion, texto_explicativo, periodo, ejercicio, agrup, 
                          cuenta_mayor_hfm, descripcion_cuenta_sific, cuenta_oracle, descripcion_cuenta_oracle, referencia) " &
                        "VALUES (@soc,@cta,@icp,@monto,@desc,@tex_explic,@period,@ejercicio,@agrup,@cta_mayor,@desc_cta_sific,@cta_oracle,@cta_desc_oracle,@TOP);",
                        conn, tran)
                        cmdIns.Parameters.AddWithValue("@soc", soc)
                        cmdIns.Parameters.AddWithValue("@cta", cta)
                        cmdIns.Parameters.AddWithValue("@icp", icp)
                        cmdIns.Parameters.AddWithValue("@monto", monto)
                        cmdIns.Parameters.AddWithValue("@desc", "Integración costo_ingreso_acum  " & soc & "-" & cta & "->" & Convert.ToString(rowCI("CUENTA_ORACLE")))
                        'campos adiconales para incertar registro nuevo.    
                        cmdIns.Parameters.AddWithValue("@tex_explic", Convert.ToString(rowCI("servicio_descrip")))
                        cmdIns.Parameters.AddWithValue("@period", Convert.ToString(rowCI("mes")))
                        cmdIns.Parameters.AddWithValue("@ejercicio", Convert.ToString(rowCI("year")))
                        cmdIns.Parameters.AddWithValue("@agrup", Convert.ToString(rowCI("clasifc_cost_ingreso")))
                        cmdIns.Parameters.AddWithValue("@cta_mayor", Convert.ToString(rowCI("CUENTA_SAP")))
                        cmdIns.Parameters.AddWithValue("@desc_cta_sific", Convert.ToString(rowCI("descrip_cuenta_sap")))
                        cmdIns.Parameters.AddWithValue("@cta_oracle", Convert.ToString(rowCI("CUENTA_ORACLE")))
                        cmdIns.Parameters.AddWithValue("@cta_desc_oracle", Convert.ToString(rowCI("descrip_cuenta_oracle")))
                        cmdIns.Parameters.AddWithValue("@TOP", "TOP=" & Convert.ToString(rowCI("TOP")))
                        cmdIns.ExecuteNonQuery()
                    End Using

                    ' 5.3) Restar del ICP None
                    Dim restante = baseSaldo - monto
                    Using cmdUpdBase As New SQLiteCommand(
                        "UPDATE t_in_sap SET saldo_acum=@rest WHERE ID=@rid;",
                        conn, tran)
                        cmdUpdBase.Parameters.AddWithValue("@rest", restante)
                        cmdUpdBase.Parameters.AddWithValue("@rid", baseId)
                        cmdUpdBase.ExecuteNonQuery()
                    End Using

                    ' Eliminar fila procesada de la tabla en memoria
                    dtCI.Rows.RemoveAt(0)
                End While

                tran.Commit()
            End Using
        End Using

        MessageBox.Show("Integración completada.", "Éxito",
                        MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ''' <summary>
    ''' Elimina ceros a la izquierda de un string.
    ''' </summary>
    Private Function TrimLeadingZeros(s As String) As String
        If s Is Nothing Then Return s
        Return s.TrimStart("0"c)
    End Function

End Class
