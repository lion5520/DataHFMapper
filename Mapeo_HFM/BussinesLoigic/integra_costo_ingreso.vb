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
                Using cmdGr As New SQLiteCommand(
                        "SELECT DISTINCT g.GRUPO 
                        FROM costo_ingreso_acum ci JOIN GL_ICP_Grupos g 
                            On LTRIM(ci.SOC_SAP,'0')=LTRIM(g.GL_ICP,'0') ORDER BY g.GRUPO;", conn, tran)
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
                Dim prompt = "Grupos disponibles:  " & String.Join(", ", grupos) & vbCrLf & "Ingrese el grupo a procesar:"
                Dim grupoSel = InputBox(prompt, "Seleccione Grupo").Trim()
                If String.IsNullOrWhiteSpace(grupoSel) OrElse Not grupos.Contains(grupoSel) Then
                    MessageBox.Show("Grupo inválido o cancelado.", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Return
                End If

                ' Paso 3: obtener sociedades para el grupo
                Dim socList As New List(Of String)
                Using cmdSoc As New SQLiteCommand(
                        "SELECT DISTINCT LTRIM(ci.SOC_SAP,'0') 
                        FROM costo_ingreso_acum ci 
                        JOIN GL_ICP_Grupos g 
                            On LTRIM(ci.SOC_SAP,'0')=LTRIM(g.GL_ICP,'0') 
                        WHERE g.GRUPO=@grp;", conn, tran)
                    cmdSoc.Parameters.AddWithValue("@grp", grupoSel)
                    Using rdr = cmdSoc.ExecuteReader()
                        While rdr.Read()
                            socList.Add(rdr.GetString(0))
                        End While
                    End Using
                End Using

                If socList.Count = 0 Then
                    MessageBox.Show("No hay sociedades asociadas al grupo.", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                ' Paso 4: cargar filas de costo_ingreso_acum para esas sociedades
                Dim dtCI As New DataTable()
                Dim inClause = String.Join(",", socList.Select(Function(s) $"'{s}'"))
                Dim sqlCI = $"SELECT * FROM costo_ingreso_acum WHERE LTRIM(SOC_SAP,'0') IN ({inClause});"
                Using daCI As New SQLiteDataAdapter(sqlCI, conn)
                    daCI.Fill(dtCI)
                End Using

                ' Obtener columnas de t_in_sap para clonación (excluir ID y rowid)
                Dim cols As New List(Of String)
                Using cmdCols As New SQLiteCommand("PRAGMA table_info(t_in_sap);", conn, tran)
                    Using rdr = cmdCols.ExecuteReader()
                        While rdr.Read()
                            Dim name = rdr("name").ToString()
                            If Not String.Equals(name, "id", StringComparison.OrdinalIgnoreCase) _
               AndAlso Not String.Equals(name, "rowid", StringComparison.OrdinalIgnoreCase) Then
                                cols.Add(name)
                            End If
                        End While
                    End Using
                End Using


                ' Lista de IDs procesados para evitar reprocesar
                Dim processedIds As New List(Of Long)

                ' Paso 5: procesar cada fila (While para poder remover)
                Dim index = 0
                While index < dtCI.Rows.Count
                    Dim rowCI = dtCI.Rows(index)
                    Dim soc = TrimLeadingZeros(Convert.ToString(rowCI("SOC_SAP")))
                    Dim cta = TrimLeadingZeros(Convert.ToString(rowCI("CUENTA_SAP")))
                    Dim icp = TrimLeadingZeros(Convert.ToString(rowCI("ICP_ORACLE")))
                    Dim monto = Math.Round(Convert.ToDouble(rowCI("MONTO_T_MXP")), 2, MidpointRounding.AwayFromZero)

                    ' Buscar coincidencia exacta en t_in_sap
                    Dim dtExist As New DataTable()
                    Using cmdE As New SQLiteCommand(
                        "SELECT ID, saldo_acum 
                        FROM t_in_sap 
                        WHERE LTRIM(sociedad,'0')=@soc AND numero_cuenta=@cta AND LTRIM(deudor_acreedor_2,'0')=@icp;", conn, tran)
                        cmdE.Parameters.AddWithValue("@soc", soc)
                        cmdE.Parameters.AddWithValue("@cta", cta)
                        cmdE.Parameters.AddWithValue("@icp", icp)
                        Using daE As New SQLiteDataAdapter(cmdE)
                            daE.Fill(dtExist)
                        End Using
                    End Using

                    Dim idExist As Long = 0, saldoExist As Double = 0
                    If dtExist.Rows.Count > 0 Then
                        idExist = dtExist.Rows(0).Field(Of Long)("ID")
                        saldoExist = Convert.ToDouble(dtExist.Rows(0)("saldo_acum"))

                        ' Si ya procesado, omitir
                        If processedIds.Contains(idExist) Then
                            index += 1 : Continue While
                        End If

                        ' Si monto igual, sólo marcar como procesado
                        If Math.Abs(saldoExist - monto) < 0.005 Then
                            processedIds.Add(idExist)
                            index += 1 : Continue While
                        End If

                        Dim diff = monto - saldoExist
                        ' Saldos difieren: actualizar y distribuir diferencia
                        Using cmdUpd As New SQLiteCommand("UPDATE t_in_sap SET saldo_acum=@new WHERE ID=@rid;", conn, tran)
                            cmdUpd.Parameters.AddWithValue("@new", monto)
                            cmdUpd.Parameters.AddWithValue("@rid", idExist)
                            cmdUpd.ExecuteNonQuery()
                        End Using


                        ' Ajustar o crear registro ICP None
                        Dim dtNone As New DataTable()
                        Using cmdN As New SQLiteCommand(
                            "SELECT ID, saldo_acum 
                            FROM t_in_sap 
                            WHERE LTRIM(sociedad,'0')=@soc AND numero_cuenta=@cta AND deudor_acreedor_2='[ICP None]';", conn, tran)
                            cmdN.Parameters.AddWithValue("@soc", soc)
                            cmdN.Parameters.AddWithValue("@cta", cta)
                            Using daN As New SQLiteDataAdapter(cmdN)
                                daN.Fill(dtNone)
                            End Using
                        End Using
                        If dtNone.Rows.Count > 0 Then
                            Dim idNone = dtNone.Rows(0).Field(Of Long)("ID")
                            Dim newNone = dtNone.Rows(0).Field(Of Double)("saldo_acum") + diff
                            Using cmdUpd2 As New SQLiteCommand("UPDATE t_in_sap SET saldo_acum=@s WHERE ID=@rid;", conn, tran)
                                cmdUpd2.Parameters.AddWithValue("@s", newNone)
                                cmdUpd2.Parameters.AddWithValue("@rid", idNone)
                                cmdUpd2.ExecuteNonQuery()
                            End Using
                        Else
                            ' Clonar original y crear nuevo
                            Dim colsList = String.Join(",", cols)
                            Dim paramList = String.Join(",", cols.Select(Function(c) "@" & c))
                            Dim sqlIns = $"INSERT INTO t_in_sap ({colsList}) VALUES ({paramList});"
                            Using cmdIns As New SQLiteCommand(sqlIns, conn, tran)
                                Dim dtOrig As New DataTable()
                                Using cmdO As New SQLiteCommand("SELECT * FROM t_in_sap WHERE ID=@rid;", conn, tran)
                                    cmdO.Parameters.AddWithValue("@rid", idExist)
                                    Using daO As New SQLiteDataAdapter(cmdO)
                                        daO.Fill(dtOrig)
                                    End Using
                                End Using
                                Dim orig = dtOrig.Rows(0)
                                For Each col In cols
                                    cmdIns.Parameters.AddWithValue("@" & col, orig(col))
                                Next
                                cmdIns.Parameters.AddWithValue("@deudor_acreedor_2", "[ICP None]")
                                cmdIns.Parameters.AddWithValue("@saldo_acum", diff)
                                cmdIns.Parameters.AddWithValue("@descripcion_cuenta_sific",
                                    "Diferencia generada por integración costo_ingreso_acum")
                                cmdIns.ExecuteNonQuery()
                            End Using
                        End If

                        processedIds.Add(idExist)
                    Else
                        ' No existe: insertar nuevo registro
                        Using cmdNew As New SQLiteCommand(
    "INSERT INTO t_in_sap (sociedad, numero_cuenta, deudor_acreedor_2, saldo_acum, descripcion_cuenta_sific) VALUES (@soc,@cta,@icp,@monto,@desc);",
    conn, tran)
                            cmdNew.Parameters.AddWithValue("@soc", soc)
                            cmdNew.Parameters.AddWithValue("@cta", cta)
                            cmdNew.Parameters.AddWithValue("@icp", icp)
                            cmdNew.Parameters.AddWithValue("@monto", monto)
                            cmdNew.Parameters.AddWithValue("@desc", "Creado por integración costo_ingreso_acum")
                            cmdNew.ExecuteNonQuery()
                            processedIds.Add(conn.LastInsertRowId)
                        End Using
                    End If
                    ' Remover de la tabla en memoria
                    dtCI.Rows.RemoveAt(index)
                End While

                tran.Commit()
            End Using
        End Using

        MessageBox.Show("Integración completada.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ''' <summary>
    ''' Elimina ceros a la izquierda de un string.
    ''' </summary>
    Private Function TrimLeadingZeros(s As String) As String
        If s Is Nothing Then Return s
        Return s.TrimStart("0"c)
    End Function
End Class
