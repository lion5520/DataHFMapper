Imports System.Data
Imports OfficeOpenXml
Imports ClosedXML.Excel
Public Class FrmPolizasHFM
    Private Sub FrmPolizasHFM_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Private repo As PolizasRepository
    Private rutaDB As String = "C:\Users\mario\Outliers\IZZI\Pruebas\mapeo_sap_sific.sqlite"

    ' Controles
    Private menuStrip As New MenuStrip()
    'Private groupApp As New GroupBox() mcl
    'Private txtUsuario, txtFecha, txtHora, txtScenario, txtYear, txtPeriod, txtValue, txtCurrDate As New TextBox()
    'Private dgvPolizasHFM, dgvCuentaOrigen, dgvReclasificar As New DataGridView()

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

        ' Grupo de aplicación
        ' groupApp.Text = "Aplicación: PRDIZZI"
        'groupApp.Size = New Size(1150, 80)
        'groupApp.Location = New Point(20, 40)
        'Dim labels = {"Usuario", "Fecha", "Hora", "Scenario", "Year", "Period", "Value", "currDate"}
        'Dim txts = {txtUsuario, txtFecha, txtHora, txtScenario, txtYear, txtPeriod, txtValue, txtCurrDate}
        'For i = 0 To labels.Length - 1
        'Dim lbl = New Label() With {.Text = labels(i), .Location = New Point(10 + i * 140, 25), .Width = 60}
        'txts(i).Location = New Point(70 + i * 140, 20)
        'txts(i).Width = 70
        'txts(i).ReadOnly = True
        '   groupApp.Controls.Add(lbl)
        '  groupApp.Controls.Add(txts(i))
        'Next
        'Me.Controls.Add(groupApp)

        ' Grids
        'dgvPolizasHFM.Location = New Point(20, 130)
        'dgvPolizasHFM.Size = New Size(1150, 350)
        dgvPolizasHFM.AllowUserToAddRows = True
        dgvPolizasHFM.AllowUserToDeleteRows = True
        dgvPolizasHFM.EditMode = DataGridViewEditMode.EditOnEnter
        dgvPolizasHFM.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        AddHandler dgvPolizasHFM.CellValueChanged, AddressOf dgvPolizasHFM_CellValueChanged
        AddHandler dgvPolizasHFM.SelectionChanged, AddressOf dgvPolizasHFM_SelectionChanged
        Me.Controls.Add(dgvPolizasHFM)

        'dgvCuentaOrigen.Location = New Point(20, 500)
        'dgvCuentaOrigen.Size = New Size(400, 80)
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
        For Each row As DataRow In dt.Rows
            ' row("Semaforo") = CalcularSemaforo(row)  mcl revisar si se habilita
        Next
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
        Dim registros = dtAll.Select($"Grupo = '{grupo.Replace("'", "''")}' AND Entity = '{entity.Replace("'", "''")}' AND Descripcion = '{descripcion.Replace("'", "''")}'", "id ASC")

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

            If Math.Abs(suma) = sumaObjetivo Then
                encontrado = True
                Exit For
                'ElseIf Math.Abs(suma) > sumaObjetivo Then
                ' Exit For
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
        If result = DialogResult.No Then Exit Sub
        repo.LimpiarTablas()
        Dim ofd As New OpenFileDialog() With {.Filter = "Archivos Excel|*.xlsx", .Title = "Selecciona el archivo Excel"}
        If ofd.ShowDialog() <> DialogResult.OK Then Exit Sub
        Try
            Using workbook As New XLWorkbook(ofd.FileName)
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
        Catch ex As Exception
            MessageBox.Show("Error al cargar archivo: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
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
End Class