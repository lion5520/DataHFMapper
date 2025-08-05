Imports System.Data
Imports System.IO
Imports ClosedXML.Excel
Imports OfficeOpenXml

Imports Workbook = ClosedXML.Excel.XLWorkbook
Public Class FrmPolizasHFM
    Private Sub FrmPolizasHFM_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Private repo As PolizasRepository
    Private rutaDB As String = "C:\Users\mario\Outliers\IZZI\Pruebas\mapeo_sap_sific.sqlite"

    ' Controles
    Private menuStrip As New MenuStrip()

    Public Sub New()
        Me.Text = "Pólizas HFM"
        Me.Size = New Size(1200, 800)
        Me.StartPosition = FormStartPosition.CenterScreen
        repo = New PolizasRepository(rutaDB)
        InitializeComponent()
        InicializarControles()
    End Sub

    Private Sub InicializarControles()
        ' Menú
        Dim mnuArchivo = New ToolStripMenuItem("Archivo")
        Dim mnuCargar = New ToolStripMenuItem("Cargar Archivo")
        Dim mnuRespaldar = New ToolStripMenuItem("Respaldar Cambios") With {.ToolTipText = "Guarda los cambios en una tabla histórica"}
        Dim mnuRecuperar = New ToolStripMenuItem("Recuperar Histórico")
        Dim mnuSalir = New ToolStripMenuItem("Salir")
        AddHandler mnuCargar.Click, AddressOf CargarArchivo_Click
        AddHandler mnuRespaldar.Click, AddressOf Respaldar_Click
        AddHandler mnuRecuperar.Click, AddressOf Recuperar_Click
        AddHandler mnuSalir.Click, Sub() Me.Close()
        mnuArchivo.DropDownItems.AddRange({mnuCargar, mnuRespaldar, mnuRecuperar, mnuSalir})
        menuStrip.Items.Add(mnuArchivo)
        Me.Controls.Add(menuStrip)


        ' Grids
        dgvPolizasHFM.AllowUserToAddRows = True
        dgvPolizasHFM.AllowUserToDeleteRows = True
        dgvPolizasHFM.EditMode = DataGridViewEditMode.EditOnEnter
        dgvPolizasHFM.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        AddHandler dgvPolizasHFM.CellValueChanged, AddressOf dgvPolizasHFM_CellValueChanged
        AddHandler dgvPolizasHFM.SelectionChanged, AddressOf dgvPolizasHFM_SelectionChanged
        Me.Controls.Add(dgvPolizasHFM)

        dgvCuentaOrigen.ReadOnly = True
        dgvCuentaOrigen.AllowUserToAddRows = False
        dgvCuentaOrigen.AllowUserToDeleteRows = False
        dgvCuentaOrigen.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        dgvCuentaOrigen.RowHeadersVisible = False
        dgvCuentaOrigen.ColumnHeadersVisible = True
        dgvCuentaOrigen.Name = "Cuenta Origen"
        Me.Controls.Add(dgvCuentaOrigen)

        'dgvReclasificar.Location = New Point(450, 500)
        'dgvReclasificar.Size = New Size(720, 120)
        dgvReclasificar.ReadOnly = True
        dgvReclasificar.AllowUserToAddRows = False
        dgvReclasificar.AllowUserToDeleteRows = False
        dgvReclasificar.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        dgvReclasificar.RowHeadersVisible = False
        dgvReclasificar.ColumnHeadersVisible = True
        dgvReclasificar.Name = "Reclasificar"
        Me.Controls.Add(dgvReclasificar)
    End Sub

    Protected Overrides Sub OnLoad(e As EventArgs)
        MyBase.OnLoad(e)
        Try
            PoblarEncabezado()
            PoblarGrids()
            PoblarCombos()
        Catch ex As Exception
            MessageBox.Show("Error al cargar datos: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub PoblarEncabezado()
        Dim dt = repo.GetEncabezado()
        If dt.Rows.Count > 0 Then
            txtUsuario.Text = dt.Rows(0)("Usuario").ToString()
            txtFecha.Text = dt.Rows(0)("Fecha").ToString()
            txtHora.Text = dt.Rows(0)("Hora").ToString()
            txtScenario.Text = dt.Rows(0)("Scenario").ToString()
            txtYear.Text = dt.Rows(0)("Year").ToString()
            txtPeriod.Text = dt.Rows(0)("Period").ToString()
            txtValue.Text = dt.Rows(0)("Value").ToString()
            txtCurrDate.Text = dt.Rows(0)("currDate").ToString()
        Else
            txtUsuario.Text = ""
            txtFecha.Text = ""
            txtHora.Text = ""
            txtScenario.Text = ""
            txtYear.Text = ""
            txtPeriod.Text = ""
            txtValue.Text = ""
            txtCurrDate.Text = ""
        End If
    End Sub

    Private Sub PoblarGrids()
        Dim dt = repo.GetPolizasHFM()
        If Not dt.Columns.Contains("Semaforo") Then
            dt.Columns.Add("Semaforo", GetType(String))
        End If
        'For Each row As DataRow In dt.Rows
        ' row("Semaforo") = CalcularSemaforo(row)  mcl revisar si se habilita
        'Next
        dgvPolizasHFM.DataSource = dt
        dgvPolizasHFM.Columns("Semaforo").DisplayIndex = 0
        dgvPolizasHFM.Columns("Semaforo").HeaderText = ""
        dgvPolizasHFM.Columns("Semaforo").Width = 30
    End Sub

    Private Function CalcularSemaforo(row As DataRow) As String
        Dim similares = repo.BuscarPolizasSimilares(row("Grupo").ToString(), row("Entity").ToString(), row("Descripcion").ToString(), CInt(row("id")))
        Dim sumaDebe = 0D
        Dim sumaHaber = 0D
        For Each r As DataRow In similares.Rows
            sumaDebe += Convert.ToDouble(r("Debe"))
            sumaHaber += Convert.ToDouble(r("Haber"))
        Next
        Dim actualDebe = Convert.ToDouble(row("Debe"))
        Dim actualHaber = Convert.ToDouble(row("Haber"))
        If similares.Rows.Count = 0 OrElse (sumaDebe + sumaHaber = 0) OrElse (sumaDebe + sumaHaber = actualDebe + actualHaber) Then
            Return ""
        Else
            Return "X"
        End If
    End Function

    Private Sub dgvPolizasHFM_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Exit Sub
        Dim row = CType(dgvPolizasHFM.Rows(e.RowIndex).DataBoundItem, DataRowView).Row
        Try
            repo.ActualizarPolizaHFM(row)
            'row("Semaforo") = CalcularSemaforo(row) mcl revisar si se habilita
            dgvPolizasHFM.InvalidateRow(e.RowIndex)
        Catch ex As Exception
            MessageBox.Show("Error al actualizar registro: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dgvPolizasHFM_SelectionChanged(sender As Object, e As EventArgs)
        Try
            If dgvPolizasHFM.SelectedRows.Count = 0 Then Exit Sub
            Dim row = CType(dgvPolizasHFM.SelectedRows(0).DataBoundItem, DataRowView).Row
            ActualizarGridsSecundarios(row)
        Catch ex As Exception
            MessageBox.Show("Error al consultar registros relacionados: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ActualizarGridsSecundarios(row As DataRow)
        dgvCuentaOrigen.DataSource = Nothing
        dgvReclasificar.DataSource = Nothing

        Dim cuentaSeleccionada = row("Account").ToString()
        Dim grupo = row("Grupo").ToString()
        Dim entity = row("Entity").ToString()
        Dim descripcion = row("Descripcion").ToString()
        Dim idSeleccionado As Integer = Convert.ToInt32(row("id"))
        Dim debeOrigen As Decimal = Convert.ToDecimal(row("Debe"))
        Dim haberOrigen As Decimal = Convert.ToDecimal(row("Haber"))
        Dim sumaOrigen As Decimal = debeOrigen - haberOrigen

        ' 2.1 Leer registros con mismo grupo, entity y descripcion, ordenados por id ascendente
        Dim dtAll = repo.GetPolizasHFM()

        ' mcl 29-jul-2025, se quita Entity del filtro porque no es necesario
        ' Dim registros = dtAll.Select($"Grupo = '{grupo.Replace("'", "''")}' AND Entity = '{entity.Replace("'", "''")}' AND Descripcion = '{descripcion.Replace("'", "''")}'", "id ASC")
        Dim registros = dtAll.Select($"Grupo = '{grupo.Replace("'", "''")}' AND Descripcion = '{descripcion.Replace("'", "''")}'", "id ASC")

        ' 2.2 Poner el registro seleccionado en dgvCuentaOrigen
        Dim origenDt As DataTable = dtAll.Clone()
        origenDt.ImportRow(row)
        dgvCuentaOrigen.DataSource = origenDt

        ' 2.2.1 Si solo hay un registro, ponerlo en dgvReclasificar
        Dim reclasDt As DataTable = dtAll.Clone()
        If registros.Length = 1 Then
            reclasDt.ImportRow(registros(0))
            dgvReclasificar.DataSource = reclasDt
            Exit Sub
        End If

        ' 2.2.1 Si sumaOrigen es igual a cero, ponerlo en el registro seleccionado
        If sumaOrigen = 0 Then
            'reclasDt.ImportRow(registros(0))
            reclasDt.ImportRow(row)
            dgvReclasificar.DataSource = reclasDt
            Exit Sub
        End If

        ' 2.2.2 Si hay más de un registro, sumar debe y restar haber de los siguientes registros (id > seleccionado)
        Dim suma As Decimal = 0D
        Dim sumaObjetivo As Decimal = Math.Abs(sumaOrigen)
        Dim usados As New List(Of DataRow)
        Dim encontrado As Boolean = False

        For Each r As DataRow In registros
            Dim idActual As Integer = Convert.ToInt32(r("id"))
            If idActual <= idSeleccionado Then Continue For

            Dim cuentaActual = r("Account").ToString()
            If cuentaActual.Equals(cuentaSeleccionada) Then Continue For

            Dim debe As Decimal = Convert.ToDecimal(r("Debe"))
            Dim haber As Decimal = Convert.ToDecimal(r("Haber"))
            suma += (debe - haber)
            usados.Add(r)

            'Encuentra el valor que iguala la suma al objetivo
            If Math.Abs(debe) = sumaObjetivo Or Math.Abs(haber) = sumaObjetivo Then
                usados.Clear()
                usados.Add(r)
                encontrado = True
                Exit For
            End If

            If Math.Abs(suma) = sumaObjetivo Then
                encontrado = True
                Exit For
            End If
        Next

        If encontrado AndAlso usados.Count > 0 Then
            For Each r In usados
                reclasDt.ImportRow(r)
            Next
        Else
            reclasDt.ImportRow(row)
        End If

        dgvReclasificar.DataSource = reclasDt
    End Sub

    Private Sub CargarArchivo_Click(sender As Object, e As EventArgs)
        Dim result = MessageBox.Show("Esta acción eliminará todos los registros de las Pólizas HFM. ¿Desea continuar?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        If result = DialogResult.No Then
            Cursor.Current = Cursors.Default
            Exit Sub
        End If
        repo.LimpiarTablas()
        Dim ofd As New OpenFileDialog() With {.Filter = "Archivos Excel|*.xlsx", .Title = "Selecciona el archivo Excel"}
        If ofd.ShowDialog() <> DialogResult.OK Then
            Cursor.Current = Cursors.Default
            Exit Sub
        End If
        Cursor.Current = Cursors.WaitCursor
        Try
            Using workbook As New Workbook(ofd.FileName)
                Dim ws = workbook.Worksheet("Polizas")
                If ws Is Nothing Then Throw New Exception("No se encontró la hoja 'Polizas'.")
                ' Validar encabezados en línea 13
                Dim headers = {"Grupo", "Etiqueta", "Descripción", "Entity", "Account", "Creado por", "Aprobado por", "Estado", "Debe", "Haber"}
                For i = 0 To headers.Length - 1
                    If ws.Cell(13, i + 1).GetString().Trim() <> headers(i) Then
                        Throw New Exception($"Encabezado incorrecto en columna {i + 1}: se esperaba '{headers(i)}'.")
                    End If
                Next

                'Actualizamos nombre del encabezado porque en la BD no tiene el acento ni espacio
                headers(2) = "Descripcion"
                headers(5) = "Creado_por"
                headers(6) = "Aprobado_por"

                ' Leer datos de encabezado
                Dim encabezadoRow = repo.GetEncabezado().NewRow()
                encabezadoRow("Usuario") = ws.Cell(2, 1).GetString().Replace("Usuario:", "").Trim()
                encabezadoRow("Fecha") = ws.Cell(3, 1).GetString().Replace("Fecha:", "").Trim()
                encabezadoRow("Hora") = ws.Cell(4, 1).GetString().Replace("Hora:", "").Trim()
                Dim scenarioLine = ws.Cell(6, 1).GetString()
                encabezadoRow("Scenario") = ExtraerValor(scenarioLine, "Scenario:")
                encabezadoRow("Year") = ExtraerValor(scenarioLine, "Year:")
                encabezadoRow("Period") = ExtraerValor(scenarioLine, "Period:")
                encabezadoRow("Value") = scenarioLine.Substring(scenarioLine.IndexOf("Value:") + 11, scenarioLine.Length - 11 - scenarioLine.IndexOf("Value:")) '  ExtraerValor(scenarioLine, "Value:")
                encabezadoRow("currDate") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                repo.InsertarEncabezado(encabezadoRow)
                PoblarEncabezado()
                ' Leer registros desde línea 15
                Dim dt = repo.GetPolizasHFM()
                For rowIdx = 15 To ws.LastRowUsed().RowNumber()

                    ' Ignorar filas vacías
                    Dim valCelda1 = ws.Cell(rowIdx, 1).GetString().Trim()
                    Dim valCelda2 = ws.Cell(rowIdx, 2).GetString().Trim()
                    Dim valCelda3 = ws.Cell(rowIdx, 3).GetString().Trim()
                    If String.IsNullOrWhiteSpace(valCelda1) And String.IsNullOrWhiteSpace(valCelda2) And
                        String.IsNullOrWhiteSpace(valCelda3) Then
                        Continue For
                    End If


                    Dim dr = dt.NewRow()
                    For colIdx = 0 To headers.Length - 1
                        If colIdx = 8 Or colIdx = 9 Then ' Debe y Haber
                            Dim valorCelda = ws.Cell(rowIdx, colIdx + 1).GetString().Trim()
                            If String.IsNullOrWhiteSpace(valorCelda) Then
                                dr(headers(colIdx)) = 0D
                            Else
                                Dim decValue As Decimal
                                If Decimal.TryParse(valorCelda, decValue) Then
                                    dr(headers(colIdx)) = decValue
                                Else
                                    dr(headers(colIdx)) = 0D ' O puedes lanzar una excepción si prefieres
                                End If
                            End If
                        Else
                            dr(headers(colIdx)) = ws.Cell(rowIdx, colIdx + 1).GetString()
                        End If
                    Next
                    repo.InsertarPolizaHFM(dr)
                Next
                PoblarGrids()
            End Using
            MessageBox.Show("Carga de Archivo finalizado.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Error al cargar archivo: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Cursor.Current = Cursors.Default
            PoblarCombos()
        End Try
    End Sub

    Private Function ExtraerValor(linea As String, etiqueta As String) As String
        Dim idx = linea.IndexOf(etiqueta)
        If idx < 0 Then Return ""
        Dim resto = linea.Substring(idx + etiqueta.Length).Trim()
        Dim siguienteEspacio = resto.IndexOf(" ")
        If siguienteEspacio > 0 Then
            resto = resto.Substring(0, siguienteEspacio)
        End If
        Return resto
    End Function

    Private Sub Respaldar_Click(sender As Object, e As EventArgs)
        MessageBox.Show("Funcionalidad de respaldo no implementada.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub Recuperar_Click(sender As Object, e As EventArgs)
        MessageBox.Show("Funcionalidad de recuperación no implementada.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ' Evento para poblar combos al iniciar
    Private Sub PoblarCombos()
        Dim dt = repo.GetPolizasHFM()

        ' Poblar cmbGrupo
        cmbGrupo.Items.Clear()
        cmbGrupo.Items.Add("") ' Valor vacío por defecto
        Dim grupos = dt.AsEnumerable().Select(Function(r) r.Field(Of String)("Grupo").Trim()).Distinct().OrderBy(Function(x) x)
        For Each g In grupos
            If Not String.IsNullOrWhiteSpace(g) Then cmbGrupo.Items.Add(g)
        Next
        cmbGrupo.SelectedIndex = 0
        cmbGrupo.Enabled = True

        ' Poblar cmbEntity
        cmbEntity.Items.Clear()
        cmbEntity.Items.Add("")

        Dim entidades = dt.AsEnumerable().Select(Function(r) r.Field(Of String)("Entity").Trim()).Distinct().OrderBy(Function(x) x)
        For Each ent In entidades
            If Not String.IsNullOrWhiteSpace(ent) Then cmbEntity.Items.Add(ent)
        Next
        cmbEntity.SelectedIndex = 0
        cmbEntity.Enabled = False

        ' Poblar cmbDescrip
        cmbDescrip.Items.Clear()
        cmbDescrip.Items.Add("")

        Dim cuentas = dt.AsEnumerable().Select(Function(r) r.Field(Of String)("Descripcion").Trim()).Distinct().OrderBy(Function(x) x)
        For Each d In cuentas
            If Not String.IsNullOrWhiteSpace(d) Then cmbDescrip.Items.Add(d)
        Next
        cmbDescrip.SelectedIndex = 0
        cmbDescrip.Enabled = False
    End Sub

    ' Evento de cambio en cmbGrupo
    Private Sub cmbGrupo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbGrupo.SelectedIndexChanged
        Dim dt = repo.GetPolizasHFM()
        Dim grupoSel = cmbGrupo.SelectedItem?.ToString()

        If String.IsNullOrWhiteSpace(grupoSel) Then
            cmbEntity.Items.Clear()
            cmbEntity.Items.Add("")
            cmbEntity.SelectedIndex = 0
            cmbEntity.Enabled = False

            cmbDescrip.Items.Clear()
            cmbDescrip.Items.Add("")
            cmbDescrip.SelectedIndex = 0
            cmbDescrip.Enabled = False
            Return
        End If

        ' Poblar y habilitar Entity filtrado por grupo
        Dim entidades = dt.AsEnumerable().
            Where(Function(r) r.Field(Of String)("Grupo").Trim() = grupoSel).
            Select(Function(r) r.Field(Of String)("Entity").Trim()).
            Distinct().OrderBy(Function(x) x)
        cmbEntity.Items.Clear()
        cmbEntity.Items.Add("")

        For Each ent In entidades
            If Not String.IsNullOrWhiteSpace(ent) Then cmbEntity.Items.Add(ent)
        Next
        cmbEntity.SelectedIndex = 0
        cmbEntity.Enabled = True

        ' Poblar y habilitar Descripcion filtrado por grupo
        Dim cuentas = dt.AsEnumerable().
            Where(Function(r) r.Field(Of String)("Grupo").Trim() = grupoSel).
            Select(Function(r) r.Field(Of String)("Descripcion").Trim()).
            Distinct().OrderBy(Function(x) x)
        cmbDescrip.Items.Clear()
        cmbDescrip.Items.Add("")

        For Each a In cuentas
            If Not String.IsNullOrWhiteSpace(a) Then cmbDescrip.Items.Add(a)
        Next
        cmbDescrip.SelectedIndex = 0
        cmbDescrip.Enabled = True
    End Sub

    Private Sub btnQuitarFiltro_Click(sender As Object, e As EventArgs) Handles btnQuitarFiltro.Click
        Try
            ' PoblarEncabezado()
            PoblarGrids()
            PoblarCombos()
        Catch ex As Exception
            MessageBox.Show("Error al cargar datos: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Private Sub btnAplicarFiltro_Click(sender As Object, e As EventArgs) Handles btnAplicarFiltro.Click
        Try
            Dim grupoFiltro As String = If(cmbGrupo.SelectedItem IsNot Nothing, cmbGrupo.SelectedItem.ToString(), "")
            Dim entityFiltro As String = If(cmbEntity.SelectedItem IsNot Nothing, cmbEntity.SelectedItem.ToString(), "")
            Dim descripFiltro As String = If(cmbDescrip.SelectedItem IsNot Nothing, cmbDescrip.SelectedItem.ToString(), "")

            If String.IsNullOrWhiteSpace(grupoFiltro) AndAlso String.IsNullOrWhiteSpace(entityFiltro) AndAlso String.IsNullOrWhiteSpace(descripFiltro) Then
                MessageBox.Show("Debe seleccionar al menos un filtro.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim dt As DataTable = repo.GetPolizasHFMFiltrado(grupoFiltro, entityFiltro, descripFiltro)

            dgvPolizasHFM.DataSource = dt
            If dt.Columns.Contains("Semaforo") Then
                dgvPolizasHFM.Columns("Semaforo").DisplayIndex = 0
                dgvPolizasHFM.Columns("Semaforo").HeaderText = ""
                dgvPolizasHFM.Columns("Semaforo").Width = 30
            End If
        Catch ex As Exception
            MessageBox.Show("Error al aplicar filtro: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnReclasificar_Click(sender As Object, e As EventArgs) Handles btnReclasificar.Click
        Cursor.Current = Cursors.WaitCursor
        Try
            ' 1.1 Obtener todos los registros de TInSap, cargados desde SAP (vista resumida)
            Dim dtTInSap = repo.GetTInSap()
            For Each regTInSap As DataRow In dtTInSap.Rows
                Dim sociedad = regTInSap("sociedad").ToString().TrimStart("0"c)
                Dim numero_cuenta = regTInSap("numero_cuenta").ToString()
                Dim periodo = regTInSap("periodo").ToString()
                Dim ejercicio = regTInSap("ejercicio").ToString()
                Dim idTInSap = Convert.ToInt32(regTInSap("id"))
                Dim saldo_acum = Convert.ToDecimal(regTInSap("saldo_acum"))


                ' 1.1.2 Buscar el registro obtenido de t_in_sap en polizas_HFM
                Dim poliza = repo.BuscarPolizaPorSociedadCuenta(sociedad, numero_cuenta)
                If poliza Is Nothing Then Continue For

                Dim descripcion = poliza("Descripcion").ToString()
                Dim grupo = poliza("Grupo").ToString()

                ' 1.1.4 Buscar registros por grupo y descripcion
                Dim polizasGrupoDesc = repo.BuscarPolizasPorGrupoDescripcion(grupo, descripcion)

                If polizasGrupoDesc.Rows.Count = 1 Then
                    Dim idPoliza = polizasGrupoDesc.Rows(0)("id").ToString()
                    repo.ActualizarTInSapAsignacion(idTInSap, "Solo un registro para reclasificar con IDPOL = " & idPoliza & " No se realiza nada")
                    Continue For
                End If

                ' Si hay varios registros, aplicar la lógica de reclasificación
                Dim sumaOrigen As Decimal
                Dim suma As Decimal = 0D
                Dim sumaObjetivo As Decimal
                Dim usados As New List(Of DataRow)
                Dim encontrado As Boolean = False
                Dim idSeleccionado As Integer
                Dim ciclo As Integer = 0
                Dim cuentaOrigen As String
                Dim sociedadOrigen As String

                For Each r As DataRow In polizasGrupoDesc.Rows

                    ' posicionarse en el registro de origen (los número de cuenta son iguales)
                    cuentaOrigen = r("Account").ToString()
                    sociedadOrigen = r("Entity").ToString()
                    If (Not cuentaOrigen.Equals(numero_cuenta) Or
                        Not sociedadOrigen.Equals(sociedad)) And
                       ciclo = 0 Then Continue For

                    ' Se utiliza Saldo como status, valor diferente de null indica que ya se proceso
                    If r.Table.Columns.Contains("saldo") Then
                        If IsDBNull(r("saldo")) OrElse r("saldo") Is Nothing Then
                            ' Si es null, se puede continuar                             
                        Else
                            ' Aquí puedes procesar r("saldo") con seguridad
                            Dim saldo As Decimal = Convert.ToDecimal(r("saldo"))
                            Continue For ' O asigna un valor: Dim saldo As Decimal = 0D
                        End If
                    End If

                    Dim idActual As Integer = Convert.ToInt32(r("id"))
                    Dim debe As Decimal = Convert.ToDecimal(r("Debe"))
                    Dim haber As Decimal = Convert.ToDecimal(r("Haber"))

                    If (ciclo = 0) Then
                        idSeleccionado = idActual
                        sumaOrigen = If(debe <> 0, -debe, haber)
                        sumaObjetivo = Math.Abs(sumaOrigen)
                        If sumaObjetivo = 0 Then Continue For
                    End If
                    ciclo += 1

                    ' Si el idActual es menor o igual al idSeleccionado, se salta este registro
                    If idActual <= idSeleccionado Then Continue For

                    ' mcl revisar
                    If (cuentaOrigen.Equals(numero_cuenta) And
                    sociedadOrigen.Equals(sociedad) And Math.Abs(sumaOrigen) = sumaObjetivo) Then Continue For

                    suma += (debe - haber)
                    usados.Add(r)

                    If Math.Abs(debe) = sumaObjetivo Or Math.Abs(haber) = sumaObjetivo Then
                        usados.Clear()
                        usados.Add(r)
                        encontrado = True
                        Exit For
                    End If

                    If Math.Abs(suma) = sumaObjetivo Then
                        encontrado = True
                        Exit For
                    End If
                Next

                ' Se modifica 4-ago-2025 
                ' Dim saldoActualizado As Decimal = If(saldo_acum < 0, saldo_acum + sumaObjetivo, saldo_acum - sumaObjetivo)
                Dim saldoActualizado As Decimal = saldo_acum + sumaObjetivo
                If encontrado AndAlso usados.Count > 0 Then
                    For Each r In usados
                        Dim sociedadHFM = r("Entity").ToString()
                        If Not sociedad.Equals(sociedadHFM) Then
                            ' revisar, ¿que se hace con la ICP? ... o ¿es correcta la asignación?
                        End If
                        Dim cuentaHFM = r("Account").ToString()
                        Dim cuentaMayorHFM = r("Account").ToString()
                        Dim debe = Convert.ToDecimal(r("Debe"))
                        Dim haber = Convert.ToDecimal(r("Haber"))
                        Dim saldoReclas As Decimal = If(debe <> 0, -debe, haber)
                        Dim idPoliza = r("id").ToString()
                        repo.InsertarTInSap(sociedadHFM, cuentaHFM, cuentaMayorHFM, saldoReclas, periodo, ejercicio, "Reclasificado de IDPOL = " & idPoliza & ", Derivado de IDTINSAP = " & idTInSap)
                        repo.ActualizarSaldoPolizaHFM(idPoliza, idSeleccionado)
                    Next
                    repo.ActualizarTInSapAsignacion(idTInSap, "Reclasificado, valor del saldo_acum anterior = " & saldo_acum, saldoActualizado)
                    ' saldo se utiliza como un status
                    repo.ActualizarSaldoPolizaHFM(idSeleccionado, idTInSap * (-1))
                Else
                    repo.ActualizarTInSapAsignacion(idTInSap, "Error en la Reclasificación, no se actualiza este registro " & saldo_acum, saldo_acum)
                End If
            Next
            MessageBox.Show("Proceso de reclasificación finalizado.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Error en reclasificación: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub

    Public Sub CrearArchivoSumariaValidacion(rutaCarpeta As String)
        Dim archivoOriginal As String = Path.Combine(rutaCarpeta, "SumariaPlantilla.xlsx")
        If Not File.Exists(archivoOriginal) Then
            Throw New FileNotFoundException("No se encontró el archivo SumariaPlantilla.xlsx en la ruta especificada.", archivoOriginal)
        End If
        Dim carpetaReportes As String = "Reportes"
        If Not Directory.Exists(carpetaReportes) Then
            Directory.CreateDirectory(carpetaReportes)
        End If

        Dim fechaActual As DateTime = DateTime.Now
        Dim nombreCopia As String = String.Format("Sumaria_{0:yyyyMMddHHmmss}.xlsx", fechaActual)
        archivoSumaria = Path.Combine(rutaCarpeta & "\" & carpetaReportes, nombreCopia)

        File.Copy(archivoOriginal, archivoSumaria, True)
    End Sub

    Dim archivoSumaria As String
    Private Sub BtnSumaria_Click(sender As Object, e As EventArgs) Handles BtnSumaria.Click
        Try
            Dim rutaActual As String = Environment.CurrentDirectory
            CrearArchivoSumariaValidacion(rutaActual)
            GenerarReporteSumariaValidacion()
        Catch ex As FileNotFoundException
            MessageBox.Show("No se encontró el archivo: " & ex.FileName, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show("Error al crear el archivo: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub GenerarReporteSumariaValidacion()
        Try
            If String.IsNullOrWhiteSpace(archivoSumaria) OrElse Not File.Exists(archivoSumaria) Then
                MessageBox.Show("No se encontró el archivo de sumaria.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim dt As DataTable = repo.GetSumariaValidacionData()
            If dt.Rows.Count = 0 Then
                MessageBox.Show("No hay datos para el reporte.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If

            Using wb As New ClosedXML.Excel.XLWorkbook(archivoSumaria)
                Dim ws = wb.Worksheet(1)
                Dim rowExcel As Integer = 10

                ' Acumuladores
                Dim acumCias As Decimal = 0, acumReclas As Decimal = 0, acumElim As Decimal = 0, acumFinal As Decimal = 0
                Dim acumCuentaCias As Decimal = 0, acumCuentaReclas As Decimal = 0, acumCuentaElim As Decimal = 0, acumCuentaFinal As Decimal = 0
                Dim acumAgrupCias As Decimal = 0, acumAgrupReclas As Decimal = 0, acumAgrupElim As Decimal = 0, acumAgrupFinal As Decimal = 0

                Dim lastCuentaOracle As String = Nothing
                Dim lastAgrupador As String = Nothing

                For i = 0 To dt.Rows.Count - 1
                    Dim r = dt.Rows(i)
                    Dim agrupador = If(IsDBNull(r("agrupador_detalle")), "", r("agrupador_detalle").ToString())
                    Dim cuentaOracle = If(IsDBNull(r("cuenta_oracle")), "", r("cuenta_oracle").ToString())
                    Dim descCuentaOracle = If(IsDBNull(r("descripcion_cuenta_oracle")), "", r("descripcion_cuenta_oracle").ToString())
                    Dim numeroCuenta = If(IsDBNull(r("numero_cuenta")), "", r("numero_cuenta").ToString())
                    Dim textoExplicativo = If(IsDBNull(r("texto_explicativo")), "", r("texto_explicativo").ToString())
                    Dim sumaCias = If(IsDBNull(r("suma_cias")), 0D, Convert.ToDecimal(r("suma_cias")))
                    Dim reclasificacion = If(IsDBNull(r("reclasificacion")), 0D, Convert.ToDecimal(r("reclasificacion")))
                    Dim eliminacion = If(IsDBNull(r("eliminacion")), 0D, Convert.ToDecimal(r("eliminacion")))
                    Dim saldoFinal = If(IsDBNull(r("saldo_final")), 0D, Convert.ToDecimal(r("saldo_final")))

                    ' Corte de cuenta_oracle
                    If lastCuentaOracle IsNot Nothing AndAlso cuentaOracle <> lastCuentaOracle Then
                        ' Línea de corte de cuenta_oracle en negritas
                        ws.Cell(rowExcel, 1).Value = lastCuentaOracle & " " & dt.Rows(i - 1)("descripcion_cuenta_oracle").ToString()
                        ws.Range(ws.Cell(rowExcel, 1), ws.Cell(rowExcel, 14)).Style.Font.Bold = True
                        ws.Cell(rowExcel, 11).Value = acumCuentaCias
                        ws.Cell(rowExcel, 12).Value = acumCuentaReclas
                        ws.Cell(rowExcel, 13).Value = acumCuentaElim
                        ws.Cell(rowExcel, 14).Value = acumCuentaFinal
                        ws.Cell(rowExcel, 11).Style.NumberFormat.Format = "#,##0.00"
                        ws.Cell(rowExcel, 12).Style.NumberFormat.Format = "#,##0.00"
                        ws.Cell(rowExcel, 13).Style.NumberFormat.Format = "#,##0.00"
                        ws.Cell(rowExcel, 14).Style.NumberFormat.Format = "#,##0.00"
                        rowExcel += 1
                        acumCuentaCias = 0 : acumCuentaReclas = 0 : acumCuentaElim = 0 : acumCuentaFinal = 0
                        ws.Row(rowExcel).Clear()
                        rowExcel += 1
                    End If

                    ' Corte de agrupador_detalle
                    If lastAgrupador IsNot Nothing AndAlso agrupador <> lastAgrupador Then
                        ws.Cell(rowExcel, 1).Value = lastAgrupador
                        ws.Range(ws.Cell(rowExcel, 2), ws.Cell(rowExcel, 14)).Style.Border.TopBorder = XLBorderStyleValues.Thin
                        ws.Range(ws.Cell(rowExcel, 1), ws.Cell(rowExcel, 14)).Style.Font.Bold = True
                        ws.Cell(rowExcel, 11).Value = acumAgrupCias
                        ws.Cell(rowExcel, 12).Value = acumAgrupReclas
                        ws.Cell(rowExcel, 13).Value = acumAgrupElim
                        ws.Cell(rowExcel, 14).Value = acumAgrupFinal
                        ws.Cell(rowExcel, 11).Style.NumberFormat.Format = "#,##0.00"
                        ws.Cell(rowExcel, 12).Style.NumberFormat.Format = "#,##0.00"
                        ws.Cell(rowExcel, 13).Style.NumberFormat.Format = "#,##0.00"
                        ws.Cell(rowExcel, 14).Style.NumberFormat.Format = "#,##0.00"
                        rowExcel += 1
                        acumAgrupCias = 0 : acumAgrupReclas = 0 : acumAgrupElim = 0 : acumAgrupFinal = 0
                        ws.Row(rowExcel).Clear()
                        ws.Row(rowExcel + 1).Clear()
                        rowExcel += 2
                    End If

                    ' Registro normal
                    ws.Cell(rowExcel, 1).Value = numeroCuenta & " " & textoExplicativo
                    ws.Cell(rowExcel, 11).Value = sumaCias
                    ws.Cell(rowExcel, 12).Value = reclasificacion
                    ws.Cell(rowExcel, 13).Value = eliminacion
                    ws.Cell(rowExcel, 14).Value = saldoFinal
                    ws.Cell(rowExcel, 11).Style.NumberFormat.Format = "#,##0.00"
                    ws.Cell(rowExcel, 12).Style.NumberFormat.Format = "#,##0.00"
                    ws.Cell(rowExcel, 13).Style.NumberFormat.Format = "#,##0.00"
                    ws.Cell(rowExcel, 14).Style.NumberFormat.Format = "#,##0.00"

                    ' Acumuladores
                    acumCias += sumaCias
                    acumReclas += reclasificacion
                    acumElim += eliminacion
                    acumFinal += saldoFinal

                    acumCuentaCias += sumaCias
                    acumCuentaReclas += reclasificacion
                    acumCuentaElim += eliminacion
                    acumCuentaFinal += saldoFinal

                    acumAgrupCias += sumaCias
                    acumAgrupReclas += reclasificacion
                    acumAgrupElim += eliminacion
                    acumAgrupFinal += saldoFinal

                    lastCuentaOracle = cuentaOracle
                    lastAgrupador = agrupador
                    rowExcel += 1
                Next

                ' Corte final de cuenta_oracle
                If dt.Rows.Count > 0 Then
                    ws.Cell(rowExcel, 1).Value = lastCuentaOracle & " " & dt.Rows(dt.Rows.Count - 1)("descripcion_cuenta_oracle").ToString()
                    ws.Range(ws.Cell(rowExcel, 1), ws.Cell(rowExcel, 14)).Style.Font.Bold = True
                    ws.Cell(rowExcel, 11).Value = acumCuentaCias
                    ws.Cell(rowExcel, 12).Value = acumCuentaReclas
                    ws.Cell(rowExcel, 13).Value = acumCuentaElim
                    ws.Cell(rowExcel, 14).Value = acumCuentaFinal
                    ws.Cell(rowExcel, 11).Style.NumberFormat.Format = "#,##0.00"
                    ws.Cell(rowExcel, 12).Style.NumberFormat.Format = "#,##0.00"
                    ws.Cell(rowExcel, 13).Style.NumberFormat.Format = "#,##0.00"
                    ws.Cell(rowExcel, 14).Style.NumberFormat.Format = "#,##0.00"
                    rowExcel += 1
                    ws.Row(rowExcel).Clear()
                    rowExcel += 1

                    ' Corte final de agrupador_detalle
                    ws.Cell(rowExcel, 1).Value = lastAgrupador
                    ws.Range(ws.Cell(rowExcel, 2), ws.Cell(rowExcel, 14)).Style.Border.TopBorder = XLBorderStyleValues.Thin
                    ws.Range(ws.Cell(rowExcel, 1), ws.Cell(rowExcel, 14)).Style.Font.Bold = True
                    ws.Cell(rowExcel, 11).Value = acumAgrupCias
                    ws.Cell(rowExcel, 12).Value = acumAgrupReclas
                    ws.Cell(rowExcel, 13).Value = acumAgrupElim
                    ws.Cell(rowExcel, 14).Value = acumAgrupFinal
                    ws.Cell(rowExcel, 11).Style.NumberFormat.Format = "#,##0.00"
                    ws.Cell(rowExcel, 12).Style.NumberFormat.Format = "#,##0.00"
                    ws.Cell(rowExcel, 13).Style.NumberFormat.Format = "#,##0.00"
                    ws.Cell(rowExcel, 14).Style.NumberFormat.Format = "#,##0.00"
                    rowExcel += 1
                    ws.Row(rowExcel).Clear()
                    ws.Row(rowExcel + 1).Clear()
                    rowExcel += 2

                    ' Corte total
                    ws.Cell(rowExcel, 1).Value = "Total"
                    ws.Range(ws.Cell(rowExcel, 2), ws.Cell(rowExcel, 14)).Style.Border.TopBorder = XLBorderStyleValues.Double
                    ws.Range(ws.Cell(rowExcel, 2), ws.Cell(rowExcel, 14)).Style.Border.BottomBorder = XLBorderStyleValues.Double
                    ws.Range(ws.Cell(rowExcel, 1), ws.Cell(rowExcel, 14)).Style.Font.Bold = True
                    ws.Cell(rowExcel, 11).Value = acumCias
                    ws.Cell(rowExcel, 12).Value = acumReclas
                    ws.Cell(rowExcel, 13).Value = acumElim
                    ws.Cell(rowExcel, 14).Value = acumFinal
                    ws.Cell(rowExcel, 11).Style.NumberFormat.Format = "#,##0.00"
                    ws.Cell(rowExcel, 12).Style.NumberFormat.Format = "#,##0.00"
                    ws.Cell(rowExcel, 13).Style.NumberFormat.Format = "#,##0.00"
                    ws.Cell(rowExcel, 14).Style.NumberFormat.Format = "#,##0.00"
                    rowExcel += 1
                    ws.Row(rowExcel).Clear()
                    ws.Row(rowExcel + 1).Clear()
                    rowExcel += 2
                End If

                wb.Save()
            End Using

            MessageBox.Show("Reporte Sumaria de Validación generado.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Error al generar el reporte: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class