Imports System.Data
Imports System.Data.SQLite
Imports System.Linq
Imports System.Windows.Forms

Public Class agrupa_cuenta_mayor
    Private ReadOnly rutaSQLite As String

    Public Sub New(ruta As String)
        Me.rutaSQLite = ruta
    End Sub

    Public Sub Procesar()
        Using conn As New SQLiteConnection($"Data Source={rutaSQLite};Version=3;")
            conn.Open()

            ' ——————————————————————————————
            ' 1) Obtener grupos disponibles, quitando ceros a la izquierda
            ' ——————————————————————————————
            Dim gruposDisponibles As New List(Of String)
            Using cmd As New SQLiteCommand(
                "SELECT DISTINCT g.GRUPO
                   FROM t_in_sap AS s
                   JOIN GL_ICP_Grupos AS g
                     ON ltrim(s.sociedad, '0') = ltrim(g.GL_ICP, '0')
                  ORDER BY g.GRUPO;", conn)
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        gruposDisponibles.Add(rdr.GetString(0).Trim())
                    End While
                End Using
            End Using

            If gruposDisponibles.Count = 0 Then
                MessageBox.Show("No se encontraron grupos para procesar.",
                                "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' ——————————————————————————————
            ' 2) Mostrar y pedir selección de GRUPO
            ' ——————————————————————————————
            MsgBox("Grupos disponibles:" & vbCrLf &
                   String.Join(vbCrLf, gruposDisponibles),
                   MsgBoxStyle.Information, "Seleccione Grupo")

            Dim grupoSel As String
            Do
                grupoSel = InputBox("Ingrese el GRUPO a procesar (o TODO para todos):",
                                    "Selección de Grupo").Trim()
                If String.Equals(grupoSel, "TODO", StringComparison.OrdinalIgnoreCase) Then
                    grupoSel = "TODO"
                    Exit Do
                End If
                Dim match = gruposDisponibles _
                    .FirstOrDefault(Function(g) String.Equals(g, grupoSel, StringComparison.OrdinalIgnoreCase))
                If match IsNot Nothing Then
                    grupoSel = match
                    Exit Do
                End If
                MsgBox("Grupo inválido. Intente de nuevo.",
                       MsgBoxStyle.Exclamation, "Error")
            Loop

            ' ——————————————————————————————
            ' 3) Llenar dtTemp aplicando mismo trim de ceros en el JOIN
            ' ——————————————————————————————
            Dim dtTemp As New DataTable()
            Using da As New SQLiteDataAdapter(
                            "SELECT
                            s.sociedad,
                            s.saldo_acum,
                            s.periodo,
                            s.ejercicio,
                            s.numero_cuenta    AS cuenta_sap,
                            s.cuenta_oracle,
                            s.cuenta_mayor_hfm,
                            s.descripcion_cuenta_sific,
                            s.deudor_acreedor_2,

                            -- Mapeo de sociedad → grupo principal
                            g1.GRUPO           AS grupo,

                            -- Mapeo de deudor_acreedor_2 → grupo secundario (ICIA_SIFIC)
                            g2.GRUPO           AS ICIA_SIFIC

                        FROM t_in_sap AS s

                        -- Join para mapa de sociedad
                        JOIN GL_ICP_Grupos AS g1
                          ON LTRIM(s.sociedad,      '0') = LTRIM(g1.GL_ICP, '0')

                        -- Join adicional para mapa de deudor_acreedor_2
                        JOIN GL_ICP_Grupos AS g2
                          ON LTRIM(s.deudor_acreedor_2, '0') = LTRIM(g2.GL_ICP, '0')

                        WHERE
                            (@grupo = 'TODO')
                            OR (g1.GRUPO = @grupo)

                        ORDER BY
                            s.periodo,
                            s.sociedad;
                        ",
                                conn)
                da.SelectCommand.Parameters.AddWithValue("@grupo", grupoSel)
                da.Fill(dtTemp)
            End Using


            If dtTemp.Rows.Count = 0 Then
                MessageBox.Show("No hay registros para el grupo seleccionado.",
                                "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' ——————————————————————————————
            ' 4) Extraer lista de periodos y construir diccionario Abrev→Periodo
            ' ——————————————————————————————
            Dim listaPeriodos = dtTemp.AsEnumerable() _
    .Select(Function(r)
                Dim tmp As Integer
                Integer.TryParse(r("periodo").ToString(), tmp)
                Return tmp
            End Function) _
    .Where(Function(n) n > 0) _
    .Distinct() _
    .OrderBy(Function(n) n) _
    .ToList()

            If listaPeriodos.Count = 0 Then
                MessageBox.Show("No se encontraron períodos para procesar.",
                    "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' Creamos un diccionario de abreviatura → número de periodo
            Dim mapaMeses As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            For Each p As Integer In listaPeriodos
                Dim abrev As String = MonthName(p, True)    ' p.ej. "Ene", "Feb", "Mar"
                mapaMeses(abrev) = p
            Next

            ' ——————————————————————————————
            ' 5) Mostrar abreviaturas y pedir selección
            ' ——————————————————————————————
            MsgBox("Períodos disponibles (abreviaturas de mes):" & vbCrLf &
       String.Join(vbCrLf, mapaMeses.Keys),
       MsgBoxStyle.Information, "Seleccione Período")

            Dim periodoInput As String
            Dim periodoProc As Integer

            Do
                periodoInput = InputBox(
        "Ingrese la abreviatura de mes a procesar (o TODO para Acumulado):",
        "Selección de Período").Trim()

                If String.Equals(periodoInput, "TODO", StringComparison.OrdinalIgnoreCase) Then
                    ' Acumulado: tomamos el último (máximo) periodo
                    periodoProc = listaPeriodos.Max()
                    Exit Do
                End If

                If mapaMeses.TryGetValue(periodoInput, periodoProc) Then
                    ' PeriodoProc ya está asignado
                    Exit Do
                End If

                MsgBox("Período inválido. Use una de las abreviaturas mostradas o TODO.",
           MsgBoxStyle.Exclamation, "Error")
            Loop

            ' A partir de aquí usa la variable periodoProc (Integer) para filtrar tu dtTemp…
            ' Por ejemplo:
            '   If periodoInput = "TODO" OrElse CInt(r("periodo")) = periodoProc Then …
            ' Y para periodo_SIFIC reutilizas:
            '   Dim mesAbrev = MonthName(periodoProc, True)


            ' ——————————————————————————————
            ' 5) Filtrar por período y agregar periodo_SIFIC
            ' ——————————————————————————————
            periodoProc = If(periodoInput = "TODO", listaPeriodos.Max(), CInt(periodoInput))
            Dim mesAbrev As String = MonthName(periodoProc, True)

            Dim dtFiltrado As DataTable = dtTemp.Clone()
            dtFiltrado.Columns.Add("periodo_SIFIC", GetType(String))

            For Each r As DataRow In dtTemp.Rows
                Dim p = CInt(r("periodo"))
                If periodoInput = "TODO" OrElse p = periodoProc Then
                    Dim nr As DataRow = dtFiltrado.NewRow()
                    For Each c As DataColumn In dtTemp.Columns
                        nr(c.ColumnName) = r(c.ColumnName)
                    Next
                    nr("periodo_SIFIC") = mesAbrev
                    dtFiltrado.Rows.Add(nr)
                End If
            Next

            ' ——————————————————————————————
            ' 6) Agrupar y sumar saldo_acum
            ' ——————————————————————————————
            Dim resultado = AgruparFilas(dtFiltrado)
            ' Creamos un nuevo DataTable con solo las filas agrupadas
            Dim dtAgrupado As DataTable = dtFiltrado.Clone()
            For Each r As DataRow In resultado
                dtAgrupado.Rows.Add(r.ItemArray)
            Next

            ' Ahora si quieres volver a usar dtFiltrado como “la tabla agrupada”:
            dtFiltrado = dtAgrupado

            ' Y si tu grid o tu inspección apuntaba a dtFiltrado, ahora verá solo las filas agrupadas

            ' ——————————————————————————————
            ' 7) Crear tabla pre_lay_out e insertar datos
            ' ——————————————————————————————
            CrearTablaSalida(conn)
            InsertarResultados(conn, resultado)

            'MsgBox("Proceso completado correctamente.",MsgBoxStyle.Information, "Éxito")
        End Using
    End Sub

    ''' <summary>
    ''' Agrupa los registros donde ICIA_SIFIC, deudor_acreedor_2 y cuenta_mayor_hfm coinciden,
    ''' sumando saldo_acum. El resto de los campos se toma de la primera fila de cada grupo.
    ''' </summary>
    Private Function AgruparFilas(dt As DataTable) As List(Of DataRow)
        Dim mapa As New Dictionary(Of String, DataRow)(StringComparer.OrdinalIgnoreCase)
        Dim salida As New List(Of DataRow)

        For Each r As DataRow In dt.Rows
            ' Generamos la llave a partir de los 3 campos
            Dim key = String.Join("|", {
            r("ICIA_SIFIC").ToString(),
            r("cuenta_oracle").ToString(),         'r("cuenta_mayor_hfm").ToString(),    Se comenta para agrupar no por cuenta mayor si no cuenta Oracle 
            r("deudor_acreedor_2").ToString()
        })

            If mapa.ContainsKey(key) Then
                ' Ya existe: sumamos saldo_acum
                mapa(key)("saldo_acum") = CDbl(mapa(key)("saldo_acum")) + CDbl(r("saldo_acum"))
            Else
                ' Nuevo grupo: clonamos la fila para la salida
                Dim nr As DataRow = dt.NewRow()
                ' Copiamos todos los campos, pero saldo_acum debe venir igual que el original
                nr("sociedad") = r("sociedad")
                nr("grupo") = r("grupo")
                nr("saldo_acum") = CDbl(r("saldo_acum"))
                nr("periodo") = r("periodo")
                nr("ejercicio") = r("ejercicio")
                nr("cuenta_sap") = r("cuenta_sap")
                nr("cuenta_oracle") = r("cuenta_oracle")
                nr("cuenta_mayor_hfm") = r("cuenta_mayor_hfm")
                nr("descripcion_cuenta_sific") = r("descripcion_cuenta_sific")
                nr("deudor_acreedor_2") = r("deudor_acreedor_2")
                nr("ICIA_SIFIC") = r("ICIA_SIFIC")
                nr("periodo_SIFIC") = r("periodo_SIFIC")
                nr("periodo_SIFIC") = r("periodo_SIFIC")
                mapa(key) = nr
                salida.Add(nr)
            End If
        Next

        Return salida
    End Function


    Private Sub CrearTablaSalida(conn As SQLiteConnection)
        Using cmd = conn.CreateCommand()
            cmd.CommandText = "DROP TABLE IF EXISTS pre_lay_out;"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "
                CREATE TABLE pre_lay_out (
                  id INTEGER PRIMARY KEY AUTOINCREMENT,
                  sociedad TEXT,
                  grupo TEXT,
                  saldo_acum REAL,
                  periodo INTEGER,
                  ejercicio INTEGER,
                  cuenta_sap TEXT,
                  cuenta_oracle TEXT,
                  cuenta_mayor_hfm TEXT,
                  descripcion_cuenta_sific TEXT,
                  deudor_acreedor_2 TEXT,
                  ICIA_SIFIC TEXT,
                  periodo_SIFIC TEXT
                );"
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub InsertarResultados(conn As SQLiteConnection, rows As List(Of DataRow))
        Using tr = conn.BeginTransaction()
            Using cmd = conn.CreateCommand()
                cmd.CommandText = "
                    INSERT INTO pre_lay_out (
                      sociedad, grupo, saldo_acum, periodo, ejercicio,
                      cuenta_sap,cuenta_oracle, cuenta_mayor_hfm, descripcion_cuenta_sific,
                      deudor_acreedor_2, ICIA_SIFIC, periodo_SIFIC
                    ) VALUES (
                      @soc, @grp, @sal, @per, @eje,
                      @cs, @co, @cm, @desc, @deu, @icia, @perS
                    );"
                cmd.Parameters.Add("@soc", DbType.String)
                cmd.Parameters.Add("@grp", DbType.String)
                cmd.Parameters.Add("@sal", DbType.Double)
                cmd.Parameters.Add("@per", DbType.Int32)
                cmd.Parameters.Add("@eje", DbType.Int32)
                cmd.Parameters.Add("@cs", DbType.String)
                cmd.Parameters.Add("@co", DbType.String)
                cmd.Parameters.Add("@cm", DbType.String)
                cmd.Parameters.Add("@desc", DbType.String)
                cmd.Parameters.Add("@deu", DbType.String)
                cmd.Parameters.Add("@icia", DbType.String)
                cmd.Parameters.Add("@perS", DbType.String)

                For Each r As DataRow In rows
                    cmd.Parameters("@soc").Value = r("sociedad")
                    cmd.Parameters("@grp").Value = r("grupo")
                    cmd.Parameters("@sal").Value = r("saldo_acum")
                    cmd.Parameters("@per").Value = r("periodo")
                    cmd.Parameters("@eje").Value = r("ejercicio")
                    cmd.Parameters("@cs").Value = r("cuenta_sap")
                    cmd.Parameters("@co").Value = r("cuenta_oracle")
                    cmd.Parameters("@cm").Value = r("cuenta_mayor_hfm")
                    cmd.Parameters("@desc").Value = r("descripcion_cuenta_sific")
                    cmd.Parameters("@deu").Value = r("deudor_acreedor_2")
                    cmd.Parameters("@icia").Value = r("ICIA_SIFIC")
                    cmd.Parameters("@perS").Value = r("periodo_SIFIC")
                    cmd.ExecuteNonQuery()
                Next
            End Using
            tr.Commit()
        End Using
    End Sub


End Class
