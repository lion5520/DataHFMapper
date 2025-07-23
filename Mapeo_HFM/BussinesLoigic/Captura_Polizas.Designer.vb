<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Captura_Polizas
    Inherits System.Windows.Forms.Form

    'Form overrides Dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    Private components As System.ComponentModel.IContainer

    '=== Controles Captura de Póliza ===
    Friend WithEvents gbCapture As System.Windows.Forms.GroupBox
    Friend WithEvents lblEscenario As System.Windows.Forms.Label
    Friend WithEvents txtEscenario As System.Windows.Forms.TextBox
    Friend WithEvents lblYear As System.Windows.Forms.Label
    Friend WithEvents txtYear As System.Windows.Forms.TextBox
    Friend WithEvents lblPeriod As System.Windows.Forms.Label
    Friend WithEvents txtPeriod As System.Windows.Forms.TextBox
    Friend WithEvents lblPeriodName As System.Windows.Forms.Label
    Friend WithEvents txtPeriodName As System.Windows.Forms.TextBox
    Friend WithEvents lblValue As System.Windows.Forms.Label
    Friend WithEvents txtValue As System.Windows.Forms.TextBox
    Friend WithEvents lblGrupo As System.Windows.Forms.Label
    Friend WithEvents txtGrupo As System.Windows.Forms.TextBox
    Friend WithEvents lblEtiqueta As System.Windows.Forms.Label
    Friend WithEvents txtEtiqueta As System.Windows.Forms.TextBox
    Friend WithEvents lblDescripcion As System.Windows.Forms.Label
    Friend WithEvents txtDescripcion As System.Windows.Forms.TextBox
    Friend WithEvents lblEntity As System.Windows.Forms.Label
    Friend WithEvents txtEntity As System.Windows.Forms.TextBox
    Friend WithEvents lblAccount As System.Windows.Forms.Label
    Friend WithEvents txtAccount As System.Windows.Forms.TextBox
    Friend WithEvents lblCreadoPor As System.Windows.Forms.Label
    Friend WithEvents txtCreadoPor As System.Windows.Forms.TextBox
    Friend WithEvents lblAprobadoPor As System.Windows.Forms.Label
    Friend WithEvents txtAprobadoPor As System.Windows.Forms.TextBox
    Friend WithEvents lblEstado As System.Windows.Forms.Label
    Friend WithEvents txtEstado As System.Windows.Forms.TextBox
    Friend WithEvents lblDebe As System.Windows.Forms.Label
    Friend WithEvents txtDebe As System.Windows.Forms.TextBox
    Friend WithEvents lblHaber As System.Windows.Forms.Label
    Friend WithEvents txtHaber As System.Windows.Forms.TextBox
    Friend WithEvents lblSaldo As System.Windows.Forms.Label
    Friend WithEvents txtSaldo As System.Windows.Forms.TextBox
    Friend WithEvents btnSave As System.Windows.Forms.Button

    '=== Controles Punto de Vista ===
    Friend WithEvents gbPuntoVista As System.Windows.Forms.GroupBox
    Friend WithEvents lblSociedad As System.Windows.Forms.Label
    Friend WithEvents cmbSociedad As System.Windows.Forms.ComboBox
    Friend WithEvents lblCuentaSAP As System.Windows.Forms.Label
    Friend WithEvents cmbCuentaSAP As System.Windows.Forms.ComboBox
    Friend WithEvents lblDeudorAcreedor2 As System.Windows.Forms.Label
    Friend WithEvents cmbDeudorAcreedor2 As System.Windows.Forms.ComboBox
    Friend WithEvents lblCuentaMayorHFM As System.Windows.Forms.Label
    Friend WithEvents cmbCuentaMayorHFM As System.Windows.Forms.ComboBox
    Friend WithEvents lblCuentaOracle As System.Windows.Forms.Label
    Friend WithEvents cmbCuentaOracle As System.Windows.Forms.ComboBox
    Friend WithEvents btnSaveSelection As System.Windows.Forms.Button
    Friend WithEvents btnResetFilters As System.Windows.Forms.Button
    Friend WithEvents dgvVista As System.Windows.Forms.DataGridView

    '=== DataGrid Pólizas ===
    Friend WithEvents dgvPolizas As System.Windows.Forms.DataGridView

    '=== Búsqueda y Exportación ===
    Friend WithEvents gbSearch As System.Windows.Forms.GroupBox
    Friend WithEvents nudSearchYear As System.Windows.Forms.NumericUpDown
    Friend WithEvents nudSearchPeriod As System.Windows.Forms.NumericUpDown
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents btnExport As System.Windows.Forms.Button

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()

        ' Instantiate controls
        Me.gbCapture = New System.Windows.Forms.GroupBox()
        Me.lblEscenario = New System.Windows.Forms.Label()
        Me.txtEscenario = New System.Windows.Forms.TextBox()
        Me.lblYear = New System.Windows.Forms.Label()
        Me.txtYear = New System.Windows.Forms.TextBox()
        Me.lblPeriod = New System.Windows.Forms.Label()
        Me.txtPeriod = New System.Windows.Forms.TextBox()
        Me.lblPeriodName = New System.Windows.Forms.Label()
        Me.txtPeriodName = New System.Windows.Forms.TextBox()
        Me.lblValue = New System.Windows.Forms.Label()
        Me.txtValue = New System.Windows.Forms.TextBox()
        Me.lblGrupo = New System.Windows.Forms.Label()
        Me.txtGrupo = New System.Windows.Forms.TextBox()
        Me.lblEtiqueta = New System.Windows.Forms.Label()
        Me.txtEtiqueta = New System.Windows.Forms.TextBox()
        Me.lblDescripcion = New System.Windows.Forms.Label()
        Me.txtDescripcion = New System.Windows.Forms.TextBox()
        Me.lblEntity = New System.Windows.Forms.Label()
        Me.txtEntity = New System.Windows.Forms.TextBox()
        Me.lblAccount = New System.Windows.Forms.Label()
        Me.txtAccount = New System.Windows.Forms.TextBox()
        Me.lblCreadoPor = New System.Windows.Forms.Label()
        Me.txtCreadoPor = New System.Windows.Forms.TextBox()
        Me.lblAprobadoPor = New System.Windows.Forms.Label()
        Me.txtAprobadoPor = New System.Windows.Forms.TextBox()
        Me.lblEstado = New System.Windows.Forms.Label()
        Me.txtEstado = New System.Windows.Forms.TextBox()
        Me.lblDebe = New System.Windows.Forms.Label()
        Me.txtDebe = New System.Windows.Forms.TextBox()
        Me.lblHaber = New System.Windows.Forms.Label()
        Me.txtHaber = New System.Windows.Forms.TextBox()
        Me.lblSaldo = New System.Windows.Forms.Label()
        Me.txtSaldo = New System.Windows.Forms.TextBox()
        Me.btnSave = New System.Windows.Forms.Button()

        Me.gbPuntoVista = New System.Windows.Forms.GroupBox()
        Me.lblSociedad = New System.Windows.Forms.Label()
        Me.cmbSociedad = New System.Windows.Forms.ComboBox()
        Me.lblCuentaSAP = New System.Windows.Forms.Label()
        Me.cmbCuentaSAP = New System.Windows.Forms.ComboBox()
        Me.lblDeudorAcreedor2 = New System.Windows.Forms.Label()
        Me.cmbDeudorAcreedor2 = New System.Windows.Forms.ComboBox()
        Me.lblCuentaMayorHFM = New System.Windows.Forms.Label()
        Me.cmbCuentaMayorHFM = New System.Windows.Forms.ComboBox()
        Me.lblCuentaOracle = New System.Windows.Forms.Label()
        Me.cmbCuentaOracle = New System.Windows.Forms.ComboBox()
        Me.btnSaveSelection = New System.Windows.Forms.Button()
        Me.btnResetFilters = New System.Windows.Forms.Button()
        Me.dgvVista = New System.Windows.Forms.DataGridView()

        Me.dgvPolizas = New System.Windows.Forms.DataGridView()

        Me.gbSearch = New System.Windows.Forms.GroupBox()
        Me.nudSearchYear = New System.Windows.Forms.NumericUpDown()
        Me.nudSearchPeriod = New System.Windows.Forms.NumericUpDown()
        Me.btnFilter = New System.Windows.Forms.Button()
        Me.btnExport = New System.Windows.Forms.Button()

        '=== gbCapture ===
        Me.gbCapture.Text = "Captura de Póliza"
        Me.gbCapture.Location = New System.Drawing.Point(12, 12)
        Me.gbCapture.Size = New System.Drawing.Size(860, 320)
        Me.gbCapture.Controls.AddRange(New Control() {
            Me.lblEscenario, Me.txtEscenario,
            Me.lblYear, Me.txtYear,
            Me.lblPeriod, Me.txtPeriod,
            Me.lblPeriodName, Me.txtPeriodName,
            Me.lblValue, Me.txtValue,
            Me.lblGrupo, Me.txtGrupo,
            Me.lblEtiqueta, Me.txtEtiqueta,
            Me.lblDescripcion, Me.txtDescripcion,
            Me.lblEntity, Me.txtEntity,
            Me.lblAccount, Me.txtAccount,
            Me.lblCreadoPor, Me.txtCreadoPor,
            Me.lblAprobadoPor, Me.txtAprobadoPor,
            Me.lblEstado, Me.txtEstado,
            Me.lblDebe, Me.txtDebe,
            Me.lblHaber, Me.txtHaber,
            Me.lblSaldo, Me.txtSaldo,
            Me.btnSave})

        ' Captura de Póliza: posiciones y textos
        Me.lblEscenario.SetBounds(10, 25, 70, 20) : Me.lblEscenario.Text = "Escenario"
        Me.txtEscenario.SetBounds(90, 25, 150, 20)
        Me.lblYear.SetBounds(260, 25, 40, 20) : Me.lblYear.Text = "Year"
        Me.txtYear.SetBounds(310, 25, 80, 20)
        Me.lblPeriod.SetBounds(405, 25, 50, 20) : Me.lblPeriod.Text = "Period"
        Me.txtPeriod.SetBounds(460, 25, 40, 20)
        Me.lblPeriodName.SetBounds(515, 25, 70, 20) : Me.lblPeriodName.Text = "Period Name"
        Me.txtPeriodName.SetBounds(595, 25, 60, 20)
        Me.lblValue.SetBounds(10, 60, 70, 20) : Me.lblValue.Text = "Value"
        Me.txtValue.SetBounds(90, 60, 150, 20)
        Me.lblGrupo.SetBounds(260, 60, 40, 20) : Me.lblGrupo.Text = "Grupo"
        Me.txtGrupo.SetBounds(310, 60, 150, 20)
        Me.lblEtiqueta.SetBounds(480, 60, 60, 20) : Me.lblEtiqueta.Text = "Etiqueta"
        Me.txtEtiqueta.SetBounds(545, 60, 150, 20)
        Me.lblDescripcion.SetBounds(10, 95, 80, 20) : Me.lblDescripcion.Text = "Descripción"
        Me.txtDescripcion.SetBounds(90, 95, 150, 20)
        Me.lblEntity.SetBounds(260, 95, 40, 20) : Me.lblEntity.Text = "Entity"
        Me.txtEntity.SetBounds(310, 95, 150, 20)
        Me.lblAccount.SetBounds(480, 95, 50, 20) : Me.lblAccount.Text = "Account"
        Me.txtAccount.SetBounds(535, 95, 150, 20)
        Me.lblCreadoPor.SetBounds(10, 130, 70, 20) : Me.lblCreadoPor.Text = "Creado por"
        Me.txtCreadoPor.SetBounds(90, 130, 150, 20)
        Me.lblAprobadoPor.SetBounds(260, 130, 80, 20) : Me.lblAprobadoPor.Text = "Aprobado por"
        Me.txtAprobadoPor.SetBounds(345, 130, 150, 20)
        Me.lblEstado.SetBounds(510, 130, 50, 20) : Me.lblEstado.Text = "Estado"
        Me.txtEstado.SetBounds(565, 130, 130, 20)
        Me.lblDebe.SetBounds(10, 165, 50, 20) : Me.lblDebe.Text = "Debe"
        Me.txtDebe.SetBounds(65, 165, 100, 20)
        Me.lblHaber.SetBounds(180, 165, 50, 20) : Me.lblHaber.Text = "Haber"
        Me.txtHaber.SetBounds(235, 165, 100, 20)
        Me.lblSaldo.SetBounds(350, 165, 50, 20) : Me.lblSaldo.Text = "Saldo"
        Me.txtSaldo.SetBounds(405, 165, 100, 20)

        Me.btnSave.SetBounds(740, 200, 100, 25)
        Me.btnSave.Text = "Agregar Póliza"

        '=== gbPuntoVista ===
        Me.gbPuntoVista.Text = "Punto de Vista"
        Me.gbPuntoVista.Location = New System.Drawing.Point(12, 350)
        Me.gbPuntoVista.Size = New System.Drawing.Size(860, 120)
        Me.gbPuntoVista.Controls.AddRange(New Control() {
            Me.lblSociedad, Me.cmbSociedad,
            Me.lblCuentaSAP, Me.cmbCuentaSAP,
            Me.lblDeudorAcreedor2, Me.cmbDeudorAcreedor2,
            Me.lblCuentaMayorHFM, Me.cmbCuentaMayorHFM,
            Me.lblCuentaOracle, Me.cmbCuentaOracle,
            Me.btnSaveSelection, Me.btnResetFilters})

        Me.lblSociedad.SetBounds(10, 25, 80, 20) : Me.lblSociedad.Text = "Sociedad"
        Me.cmbSociedad.SetBounds(95, 25, 150, 20)
        Me.lblCuentaSAP.SetBounds(260, 25, 80, 20) : Me.lblCuentaSAP.Text = "Cuenta_SAP"
        Me.cmbCuentaSAP.SetBounds(345, 25, 150, 20)
        Me.lblDeudorAcreedor2.SetBounds(510, 25, 110, 20) : Me.lblDeudorAcreedor2.Text = "Deudor/Acreedor"
        Me.cmbDeudorAcreedor2.SetBounds(625, 25, 150, 20)
        Me.lblCuentaMayorHFM.SetBounds(10, 60, 100, 20) : Me.lblCuentaMayorHFM.Text = "Cuenta Mayor HFM"
        Me.cmbCuentaMayorHFM.SetBounds(115, 60, 150, 20)
        Me.lblCuentaOracle.SetBounds(280, 60, 80, 20) : Me.lblCuentaOracle.Text = "Cuenta Oracle"
        Me.cmbCuentaOracle.SetBounds(365, 60, 150, 20)
        Me.btnSaveSelection.SetBounds(540, 60, 120, 25) : Me.btnSaveSelection.Text = "Guardar Selección"
        Me.btnResetFilters.SetBounds(680, 60, 120, 25) : Me.btnResetFilters.Text = "Reset Filtros"

        '=== dgvVista ===
        Me.dgvVista.Location = New System.Drawing.Point(12, 480)
        Me.dgvVista.Size = New System.Drawing.Size(860, 150)

        '=== dgvPolizas ===
        Me.dgvPolizas.Location = New System.Drawing.Point(12, 640)
        Me.dgvPolizas.Size = New System.Drawing.Size(860, 200)

        '=== gbSearch ===
        Me.gbSearch.Text = "Búsqueda y Exportación"
        Me.gbSearch.Location = New System.Drawing.Point(12, 850)
        Me.gbSearch.Size = New System.Drawing.Size(860, 80)
        Me.gbSearch.Controls.AddRange(New Control() {
            Me.nudSearchYear, Me.nudSearchPeriod, Me.btnFilter, Me.btnExport})

        Me.nudSearchYear.Location = New System.Drawing.Point(60, 30)
        Me.nudSearchYear.Size = New System.Drawing.Size(60, 20)
        Me.nudSearchPeriod.Location = New System.Drawing.Point(200, 30)
        Me.nudSearchPeriod.Size = New System.Drawing.Size(60, 20)
        Me.btnFilter.SetBounds(280, 30, 75, 23) : Me.btnFilter.Text = "Filtrar"
        Me.btnExport.SetBounds(775, 30, 75, 23) : Me.btnExport.Text = "Exportar"

        '=== Form ===
        Me.ClientSize = New System.Drawing.Size(884, 950)
        Me.Controls.AddRange(New Control() {
            Me.gbCapture,
            Me.gbPuntoVista,
            Me.dgvVista,
            Me.dgvPolizas,
            Me.gbSearch})
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Captura_Polizas"

        CType(Me.dgvVista, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvPolizas, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudSearchYear, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudSearchPeriod, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub
End Class
