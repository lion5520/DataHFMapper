Imports System
Imports System.Data
Imports System.Data.SQLite
Imports System.IO
Imports System.Windows.Forms
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Public Class Captura_Polizas
    Inherits Form

    Private ReadOnly dbPath As String
    Private dtPolizas As DataTable
    Private dtVista As DataTable

    Public Sub New(databasePath As String, Optional parentForm As Form = Nothing)
        InitializeComponent()

        dbPath = databasePath

        ' Autollenado: solo lectura
        For Each tb In New TextBox() {txtEscenario, txtYear, txtPeriod, txtPeriodName, txtGrupo}
            tb.ReadOnly = True
            tb.BackColor = SystemColors.ControlLight
        Next

        ' Botones iniciales
        btnSave.Enabled = False
        btnSaveSelection.Enabled = False

        ' Validaciones TextChanged
        For Each tb In New TextBox() {
            txtEtiqueta, txtDescripcion, txtEntity,
            txtAccount, txtCreadoPor, txtAprobadoPor,
            txtEstado, txtValue, txtDebe, txtHaber}
            AddHandler tb.TextChanged, AddressOf ValidateInputs
        Next

        ' Solo numérico y 2 decimales
        AddHandler txtDebe.KeyPress, AddressOf NumericOnly_KeyPress
        AddHandler txtHaber.KeyPress, AddressOf NumericOnly_KeyPress

        ' Load event
        AddHandler Me.Load, AddressOf Captura_Polizas_Load

        ' Filtros cascada
        AddHandler cmbSociedad.SelectedIndexChanged, AddressOf OnSociedadChanged
        AddHandler cmbCuentaSAP.SelectedIndexChanged, AddressOf OnCuentaSAPChanged
        AddHandler cmbDeudorAcreedor2.SelectedIndexChanged, AddressOf OnDeudorAcreedor2Changed
        AddHandler cmbCuentaMayorHFM.SelectedIndexChanged, AddressOf OnCuentaMayorHFMChanged
        AddHandler cmbCuentaOracle.SelectedIndexChanged, AddressOf OnFiltroChanged

        ' Botones
        AddHandler btnSave.Click, AddressOf BtnSave_Click
        AddHandler btnFilter.Click, AddressOf BtnFilter_Click
        AddHandler btnExport.Click, AddressOf BtnExport_Click
        AddHandler btnSaveSelection.Click, AddressOf BtnSaveSelection_Click
        AddHandler btnResetFilters.Click, AddressOf BtnResetFilters_Click

        ' Modal behavior
        If parentForm IsNot Nothing Then
            Me.Owner = parentForm
            parentForm.Enabled = False
            AddHandler Me.FormClosed, Sub() parentForm.Enabled = True
        End If
    End Sub

    Private Sub Captura_Polizas_Load(sender As Object, e As EventArgs)
        FillInitialFields()
        LoadGridData()
        LoadSociedadFilter()
    End Sub

    '— Autollenado campos iniciales —'
    Private Sub FillInitialFields()
        Using cn = New SQLiteConnection($"Data Source={dbPath};Version=3;")
            cn.Open()
            Dim sql = "
SELECT t.Periodo
     , t.sociedad
     , COALESCE(g.GRUPO,'') AS GRUPO
FROM t_in_sap AS t
LEFT JOIN GL_ICP_Grupos AS g
  ON g.GL_ICP = LTRIM(t.sociedad, '0')
LIMIT 1;"

            Using cmd = New SQLiteCommand(sql, cn), dr = cmd.ExecuteReader()
                If dr.Read() Then
                    txtEscenario.Text = "REAL"
                    txtYear.Text = DateTime.Now.Year.ToString()
                    Dim per = CInt(dr("Periodo"))
                    txtPeriod.Text = per.ToString()
                    txtPeriodName.Text = MonthName(per, True).Substring(0, 3).ToUpper()
                    txtGrupo.Text = dr("GRUPO").ToString()
                Else
                    gbCapture.Enabled = False
                End If
            End Using
        End Using
    End Sub

    '— Carga DataGrid Pólizas —'
    Private Sub LoadGridData(Optional yr As Integer? = Nothing, Optional pr As Integer? = Nothing)
        Using cn = New SQLiteConnection($"Data Source={dbPath};Version=3;")
            cn.Open()
            Dim sql = "SELECT * FROM polizas_HFM"
            If yr.HasValue AndAlso pr.HasValue Then
                sql &= " WHERE Year=@yr AND Period=@pr"
            End If
            Using da = New SQLiteDataAdapter(sql, cn)
                If sql.Contains("WHERE") Then
                    da.SelectCommand.Parameters.AddWithValue("@yr", yr.Value)
                    da.SelectCommand.Parameters.AddWithValue("@pr", pr.Value)
                End If
                dtPolizas = New DataTable()
                da.Fill(dtPolizas)
                dgvPolizas.DataSource = dtPolizas
            End Using
        End Using
    End Sub

    '— Guardar Póliza —'
    Private Sub BtnSave_Click(sender As Object, e As EventArgs)
        Using cn = New SQLiteConnection($"Data Source={dbPath};Version=3;"),
              cmd = New SQLiteCommand("
INSERT INTO polizas_HFM
(Escenario, Year, Period, Period_name, Grupo, Etiqueta, Descripcion, Entity,
 Account, Creado_por, Aprobado_por, Estado, Value, Debe, Haber, Saldo)
VALUES
(@esc,@yr,@pr,@prn,@grp,@etq,@desc,@ent,@acc,@cp,@ap,@est,@val,@deb,@hab,@sal)
", cn)
            cn.Open()
            With cmd.Parameters
                .AddWithValue("@esc", txtEscenario.Text)
                .AddWithValue("@yr", CInt(txtYear.Text))
                .AddWithValue("@pr", CInt(txtPeriod.Text))
                .AddWithValue("@prn", txtPeriodName.Text)
                .AddWithValue("@grp", txtGrupo.Text)
                .AddWithValue("@etq", txtEtiqueta.Text)
                .AddWithValue("@desc", txtDescripcion.Text)
                .AddWithValue("@ent", txtEntity.Text)
                .AddWithValue("@acc", txtAccount.Text)
                .AddWithValue("@cp", txtCreadoPor.Text)
                .AddWithValue("@ap", txtAprobadoPor.Text)
                .AddWithValue("@est", txtEstado.Text)
                .AddWithValue("@val", If(txtValue.Text.Trim = "", DBNull.Value, txtValue.Text))
                .AddWithValue("@deb", If(txtDebe.Text.Trim = "", DBNull.Value, CDec(txtDebe.Text)))
                .AddWithValue("@hab", If(txtHaber.Text.Trim = "", DBNull.Value, CDec(txtHaber.Text)))
                .AddWithValue("@sal", CDec(txtSaldo.Text))
            End With
            cmd.ExecuteNonQuery()
        End Using
        LoadGridData()
    End Sub

    '— Filtrar Pólizas —'
    Private Sub BtnFilter_Click(sender As Object, e As EventArgs)
        LoadGridData(CInt(nudSearchYear.Value), CInt(nudSearchPeriod.Value))
    End Sub

    '— Exportar a Excel —'
    Private Sub BtnExport_Click(sender As Object, e As EventArgs)
        Using dlg As New SaveFileDialog() With {.Filter = "Excel Files|*.xlsx", .FileName = "Polizas.xlsx"}
            If dlg.ShowDialog() <> DialogResult.OK Then Return
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial
            Using pkg = New ExcelPackage(), ws = pkg.Workbook.Worksheets.Add("Pólizas")
                Dim r = 1
                ws.Cells(r, 1).Value = "Aplicación: PRDIZZI" : r += 1
                ws.Cells(r, 1).Value = $"Usuario: {Environment.UserName}@IZZI" : r += 1
                ws.Cells(r, 1).Value = $"Fecha: {Now:dd/MM/yyyy}" : r += 1
                ws.Cells(r, 1).Value = $"Hora: {Now:HH:mm:ss}" : r += 2
                ' –– Encabezados y agrupado por Grupo (idéntico a tu versión previa) ––
                pkg.SaveAs(New FileInfo(dlg.FileName))
            End Using
            MessageBox.Show("Exportación completada", "Exportar", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Using
    End Sub

    '— Validaciones —'
    Private Sub ValidateInputs(sender As Object, e As EventArgs)
        Dim camposOk = {txtEtiqueta, txtDescripcion, txtEntity, txtAccount, txtCreadoPor, txtAprobadoPor, txtEstado, txtValue} _
                       .All(Function(tb) Not String.IsNullOrWhiteSpace(tb.Text))
        Dim monedasOk = Decimal.TryParse(txtDebe.Text, Nothing) AndAlso Decimal.TryParse(txtHaber.Text, Nothing)
        btnSave.Enabled = camposOk AndAlso monedasOk
    End Sub

    Private Sub NumericOnly_KeyPress(sender As Object, e As KeyPressEventArgs)
        Dim tb = DirectCast(sender, TextBox)
        If Char.IsControl(e.KeyChar) Then Return
        If Char.IsDigit(e.KeyChar) Then
            Dim parts = tb.Text.Split("."c)
            If parts.Length = 2 AndAlso tb.SelectionStart > tb.Text.IndexOf("."c) AndAlso parts(1).Length >= 2 Then
                e.Handled = True
            End If
            Return
        End If
        If e.KeyChar = "."c Then
            If tb.Text.Contains("."c) OrElse tb.SelectionStart = 0 Then e.Handled = True
            Return
        End If
        e.Handled = True
    End Sub

    '— Punto de Vista Cascada —'
    Private Sub LoadSociedadFilter()
        cmbSociedad.Items.Clear()
        Using cn = New SQLiteConnection($"Data Source={dbPath};Version=3;"),
              cmd = New SQLiteCommand("SELECT DISTINCT sociedad FROM t_in_sap ORDER BY sociedad", cn)
            cn.Open()
            Using dr = cmd.ExecuteReader()
                While dr.Read()
                    cmbSociedad.Items.Add(dr.GetString(0))
                End While
            End Using
        End Using
        LoadVistaData()
    End Sub

    Private Sub OnSociedadChanged(sender As Object, e As EventArgs) _
        Handles cmbSociedad.SelectedIndexChanged
        ' Limpiar abajo y rellenar cuenta SAP...
        cmbCuentaSAP.Items.Clear()
        Using cn = New SQLiteConnection($"Data Source={dbPath};Version=3;"),
              cmd = New SQLiteCommand(
                "SELECT DISTINCT numero_cuenta FROM t_in_sap WHERE sociedad=@soc ORDER BY numero_cuenta", cn)
            cn.Open()
            cmd.Parameters.AddWithValue("@soc", cmbSociedad.Text)
            Using dr = cmd.ExecuteReader()
                While dr.Read()
                    cmbCuentaSAP.Items.Add(dr.GetString(0))
                End While
            End Using
        End Using
        ' Reset siguientes
        cmbDeudorAcreedor2.Items.Clear()
        cmbCuentaMayorHFM.Items.Clear()
        cmbCuentaOracle.Items.Clear()
        LoadVistaData()
    End Sub

    Private Sub OnCuentaSAPChanged(sender As Object, e As EventArgs) _
        Handles cmbCuentaSAP.SelectedIndexChanged
        ' Rellenar deudor_acreedor_2 para la sociedad+cuenta
        cmbDeudorAcreedor2.Items.Clear()
        Using cn = New SQLiteConnection($"Data Source={dbPath};Version=3;"),
              cmd = New SQLiteCommand(
                "SELECT DISTINCT deudor_acreedor_2 FROM t_in_sap
                 WHERE sociedad=@soc AND numero_cuenta=@cta
                 ORDER BY deudor_acreedor_2", cn)
            cn.Open()
            cmd.Parameters.AddWithValue("@soc", cmbSociedad.Text)
            cmd.Parameters.AddWithValue("@cta", cmbCuentaSAP.Text)
            Using dr = cmd.ExecuteReader()
                While dr.Read()
                    cmbDeudorAcreedor2.Items.Add(dr.GetString(0))
                End While
            End Using
        End Using
        cmbCuentaMayorHFM.Items.Clear()
        cmbCuentaOracle.Items.Clear()
        LoadVistaData()
    End Sub

    Private Sub OnDeudorAcreedor2Changed(sender As Object, e As EventArgs) _
        Handles cmbDeudorAcreedor2.SelectedIndexChanged
        ' Rellenar cuenta_mayor_hfm
        cmbCuentaMayorHFM.Items.Clear()
        Using cn = New SQLiteConnection($"Data Source={dbPath};Version=3;"),
              cmd = New SQLiteCommand(
                "SELECT DISTINCT cuenta_mayor_hfm FROM t_in_sap
                 WHERE sociedad=@soc AND numero_cuenta=@cta AND deudor_acreedor_2=@da2
                 ORDER BY cuenta_mayor_hfm", cn)
            cn.Open()
            cmd.Parameters.AddWithValue("@soc", cmbSociedad.Text)
            cmd.Parameters.AddWithValue("@cta", cmbCuentaSAP.Text)
            cmd.Parameters.AddWithValue("@da2", cmbDeudorAcreedor2.Text)
            Using dr = cmd.ExecuteReader()
                While dr.Read()
                    cmbCuentaMayorHFM.Items.Add(dr.GetString(0))
                End While
            End Using
        End Using
        cmbCuentaOracle.Items.Clear()
        LoadVistaData()
    End Sub

    Private Sub OnCuentaMayorHFMChanged(sender As Object, e As EventArgs) _
        Handles cmbCuentaMayorHFM.SelectedIndexChanged
        ' Rellenar cuenta_oracle
        cmbCuentaOracle.Items.Clear()
        Using cn = New SQLiteConnection($"Data Source={dbPath};Version=3;"),
              cmd = New SQLiteCommand(
                "SELECT DISTINCT cuenta_oracle FROM t_in_sap
                 WHERE sociedad=@soc AND numero_cuenta=@cta
                   AND deudor_acreedor_2=@da2
                   AND cuenta_mayor_hfm=@hmf
                 ORDER BY cuenta_oracle", cn)
            cn.Open()
            cmd.Parameters.AddWithValue("@soc", cmbSociedad.Text)
            cmd.Parameters.AddWithValue("@cta", cmbCuentaSAP.Text)
            cmd.Parameters.AddWithValue("@da2", cmbDeudorAcreedor2.Text)
            cmd.Parameters.AddWithValue("@hmf", cmbCuentaMayorHFM.Text)
            Using dr = cmd.ExecuteReader()
                While dr.Read()
                    cmbCuentaOracle.Items.Add(dr.GetString(0))
                End While
            End Using
        End Using
        LoadVistaData()
    End Sub

    Private Sub OnFiltroChanged(sender As Object, e As EventArgs) _
        Handles cmbCuentaOracle.SelectedIndexChanged
        LoadVistaData()
    End Sub

    Private Sub LoadVistaData()
        Using cn = New SQLiteConnection($"Data Source={dbPath};Version=3;")
            cn.Open()
            Dim filtros As New List(Of String)
            Dim cmd As New SQLiteCommand("", cn)
            If cmbSociedad.SelectedIndex >= 0 Then
                filtros.Add("sociedad=@soc") : cmd.Parameters.AddWithValue("@soc", cmbSociedad.Text)
            End If
            If cmbCuentaSAP.SelectedIndex >= 0 Then
                filtros.Add("numero_cuenta=@cta") : cmd.Parameters.AddWithValue("@cta", cmbCuentaSAP.Text)
            End If
            If cmbDeudorAcreedor2.SelectedIndex >= 0 Then
                filtros.Add("deudor_acreedor_2=@da2") : cmd.Parameters.AddWithValue("@da2", cmbDeudorAcreedor2.Text)
            End If
            If cmbCuentaMayorHFM.SelectedIndex >= 0 Then
                filtros.Add("cuenta_mayor_hfm=@hmf") : cmd.Parameters.AddWithValue("@hmf", cmbCuentaMayorHFM.Text)
            End If
            If cmbCuentaOracle.SelectedIndex >= 0 Then
                filtros.Add("cuenta_oracle=@ora") : cmd.Parameters.AddWithValue("@ora", cmbCuentaOracle.Text)
            End If

            Dim sql = "SELECT sociedad, numero_cuenta AS Cuenta_SAP, deudor_acreedor_2,
                       cuenta_mayor_hfm, cuenta_oracle, texto_explicativo, saldo_acum
                       FROM t_in_sap"
            If filtros.Count > 0 Then sql &= " WHERE " & String.Join(" AND ", filtros)
            cmd.CommandText = sql

            dtVista = New DataTable()
            Using da = New SQLiteDataAdapter(cmd)
                da.Fill(dtVista)
            End Using
            dgvVista.DataSource = dtVista
        End Using
    End Sub

    '— Exportar selección Punto de Vista —'
    Private Sub BtnSaveSelection_Click(sender As Object, e As EventArgs)
        If dgvVista.SelectedRows.Count = 0 Then
            MessageBox.Show("Seleccione al menos una fila.", "Guardar Selección", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Using dlg As New SaveFileDialog() With {.Filter = "CSV Files|*.csv", .FileName = "Seleccion_t_in_sap.csv"}
            If dlg.ShowDialog() <> DialogResult.OK Then Return
            Using sw As New StreamWriter(dlg.FileName)
                ' encabezados
                sw.WriteLine(String.Join(",", dtVista.Columns.Cast(Of DataColumn).Select(Function(c) c.ColumnName)))
                For Each row As DataGridViewRow In dgvVista.SelectedRows
                    Dim vals = dtVista.Columns.Cast(Of DataColumn)().
                               Select(Function(c) row.Cells(c.ColumnName).Value?.ToString().Replace(",", ""))
                    sw.WriteLine(String.Join(",", vals))
                Next
            End Using
            MessageBox.Show("Exportada CSV con éxito.", "Guardar Selección", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Using
    End Sub

    '— Reset filtros Punto de Vista —'
    Private Sub BtnResetFilters_Click(sender As Object, e As EventArgs)
        cmbSociedad.SelectedIndex = -1
        cmbCuentaSAP.SelectedIndex = -1
        cmbDeudorAcreedor2.SelectedIndex = -1
        cmbCuentaMayorHFM.SelectedIndex = -1
        cmbCuentaOracle.SelectedIndex = -1
        LoadVistaData()
    End Sub

End Class
