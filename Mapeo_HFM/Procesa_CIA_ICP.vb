Imports System
Imports System.Data
Imports System.Data.SQLite
Imports System.Windows.Forms

Public Class Procesa_CIA_ICP

    ''' <summary>
    ''' Lee todos los registros de t_in_sap,
    ''' asigna CIA e ICP según GL_ICP_Grupos.GRUPO,
    ''' limpia lay_out y luego inserta los resultados con estos campos:
    '''   CIA, ICP, MON='MX', TIPO='H', YEAR, MES,
    '''   CTA = cuenta_oracle,
    '''   IMPORTE = saldo_acum,
    '''   DESCRIP = descripcion_cuenta_sific,
    '''   TOP = 0 si ICP en {'0','00','000'}, 1 en otro caso,
    '''   S_NEG = 2000.
    ''' </summary>
    ''' <param name="dbPath">Ruta al archivo .sqlite</param>
    Public Shared Sub Ejecutar(dbPath As String)
        Dim connString = $"Data Source={dbPath};Version=3;"
        Dim warnings As New List(Of String)

        Try
            Using conn As New SQLiteConnection(connString)
                conn.Open()
                Using tx = conn.BeginTransaction()

                    ' 1) Limpiar lay_out
                    Using cmdClear As New SQLiteCommand("DELETE FROM lay_out", conn, tx)
                        cmdClear.ExecuteNonQuery()
                        cmdClear.CommandText = "DELETE FROM sqlite_sequence WHERE name = 'lay_out';"
                        cmdClear.ExecuteNonQuery()
                    End Using

                    ' 2) Leer t_in_sap con los campos extra y LEFT JOIN para CIA/ICP
                    Dim sql =
                        "SELECT s.sociedad," &
                        "       s.deudor_acreedor_2," &
                        "       s.cuenta_oracle," &
                        "       s.saldo_acum," &
                        "       s.descripcion_cuenta_sific," &
                        "       g1.GRUPO AS CIA," &
                        "       g2.GRUPO AS ICP " &
                        "  FROM t_in_sap s " &
                        "  LEFT JOIN GL_ICP_Grupos g1 ON LTRIM(s.sociedad, '0')         = LTRIM(g1.GL_ICP, '0') " &
                        "  LEFT JOIN GL_ICP_Grupos g2 ON LTRIM(s.deudor_acreedor_2, '0') = LTRIM(g2.GL_ICP, '0');"

                    Dim dt As New DataTable()
                    Using cmdSel As New SQLiteCommand(sql, conn, tx)
                        Using da As New SQLiteDataAdapter(cmdSel)
                            da.Fill(dt)
                        End Using
                    End Using

                    ' 3) Preparar INSERT con todos los nuevos campos
                    Dim insertSql =
                        "INSERT INTO lay_out " &
                        "(CIA, ICP, MON, TIPO, YEAR, MES, " &
                        " CTA, IMPORTE, DESCRIP, TOP, S_NEG) " &
                        "VALUES (@cia,@icp,@mon,@tipo,@year,@mes," &
                        " @cta,@importe,@descrip,@top,@s_neg);"

                    Using cmdIns As New SQLiteCommand(insertSql, conn, tx)
                        cmdIns.Parameters.Add("@cia", DbType.String)
                        cmdIns.Parameters.Add("@icp", DbType.String)
                        cmdIns.Parameters.Add("@mon", DbType.String)
                        cmdIns.Parameters.Add("@tipo", DbType.String)
                        cmdIns.Parameters.Add("@year", DbType.Int32)
                        cmdIns.Parameters.Add("@mes", DbType.Int32)
                        cmdIns.Parameters.Add("@cta", DbType.String)
                        cmdIns.Parameters.Add("@importe", DbType.Double)
                        cmdIns.Parameters.Add("@descrip", DbType.String)
                        cmdIns.Parameters.Add("@top", DbType.Int32)
                        cmdIns.Parameters.Add("@s_neg", DbType.Int32)
                        cmdIns.Prepare()

                        ' 4) Fallback sin ceros
                        Dim fallbackSql = "SELECT GRUPO FROM GL_ICP_Grupos WHERE GL_ICP = @key"
                        Using cmdFbk As New SQLiteCommand(fallbackSql, conn, tx)
                            cmdFbk.Parameters.Add("@key", DbType.String)

                            Dim currentYear = DateTime.Now.Year
                            Dim currentMonth = DateTime.Now.Month

                            For Each row As DataRow In dt.Rows
                                Dim soc = row("sociedad").ToString()
                                Dim dea = row("deudor_acreedor_2").ToString()

                                ' Campos originales
                                Dim cia = If(row.IsNull("CIA"), String.Empty, row("CIA").ToString())
                                Dim icp = If(row.IsNull("ICP"), String.Empty, row("ICP").ToString())
                                Dim cta = If(row.IsNull("cuenta_oracle"), String.Empty, row("cuenta_oracle").ToString())
                                Dim imp = If(row.IsNull("saldo_acum"), 0D, Convert.ToDouble(row("saldo_acum")))
                                Dim desc = If(row.IsNull("descripcion_cuenta_sific"), String.Empty, row("descripcion_cuenta_sific").ToString())

                                ' Reintento CIA
                                If String.IsNullOrEmpty(cia) Then
                                    Dim keySoc = soc.TrimStart("0"c)
                                    If keySoc <> "" Then
                                        cmdFbk.Parameters("@key").Value = keySoc
                                        Dim alt = cmdFbk.ExecuteScalar()
                                        cia = If(alt IsNot Nothing, alt.ToString(), String.Empty)
                                    End If
                                End If

                                ' Reintento ICP
                                If String.IsNullOrEmpty(icp) Then
                                    Dim keyIcp = dea.TrimStart("0"c)
                                    If keyIcp <> "" Then
                                        cmdFbk.Parameters("@key").Value = keyIcp
                                        Dim alt2 = cmdFbk.ExecuteScalar()
                                        icp = If(alt2 IsNot Nothing, alt2.ToString(), String.Empty)
                                    End If
                                End If

                                ' Advertencias
                                If cia = "" Then warnings.Add($"Sociedad '{soc}' sin grupo CIA.")
                                If icp = "" Then warnings.Add($"Deudor/Acreedor '{dea}' sin grupo ICP.")

                                ' Calcular TOP
                                Dim topVal As Integer = If(icp = "0" OrElse icp = "00" OrElse icp = "000", 0, 1)

                                ' 5) Asignar parámetros y ejecutar inserción
                                cmdIns.Parameters("@cia").Value = cia
                                cmdIns.Parameters("@icp").Value = icp
                                cmdIns.Parameters("@mon").Value = "MX"
                                cmdIns.Parameters("@tipo").Value = "H"
                                cmdIns.Parameters("@year").Value = currentYear
                                cmdIns.Parameters("@mes").Value = currentMonth
                                cmdIns.Parameters("@cta").Value = cta
                                cmdIns.Parameters("@importe").Value = imp
                                cmdIns.Parameters("@descrip").Value = desc
                                cmdIns.Parameters("@top").Value = topVal
                                cmdIns.Parameters("@s_neg").Value = 2000

                                cmdIns.ExecuteNonQuery()
                            Next
                        End Using
                    End Using

                    tx.Commit()
                End Using
            End Using

            ' 6) Mostrar advertencias
            If warnings.Count > 0 Then
                MessageBox.Show(
                    String.Join(Environment.NewLine, warnings.Distinct()),
                    "Advertencias Procesa_CIA_ICP",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                MessageBox.Show(
                    "Proceso CIA/ICP completado con éxito.",
                    "Éxito",
                    MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show(
                $"Error en Procesa_CIA_ICP:{Environment.NewLine}{ex.Message}",
                "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class

