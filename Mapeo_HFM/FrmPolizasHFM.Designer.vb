<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmPolizasHFM
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.groupApp = New System.Windows.Forms.GroupBox()
        Me.txtCurrDate = New System.Windows.Forms.TextBox()
        Me.txtValue = New System.Windows.Forms.TextBox()
        Me.txtPeriod = New System.Windows.Forms.TextBox()
        Me.txtYear = New System.Windows.Forms.TextBox()
        Me.txtScenario = New System.Windows.Forms.TextBox()
        Me.txtHora = New System.Windows.Forms.TextBox()
        Me.txtFecha = New System.Windows.Forms.TextBox()
        Me.txtUsuario = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dgvPolizasHFM = New System.Windows.Forms.DataGridView()
        Me.dgvCuentaOrigen = New System.Windows.Forms.DataGridView()
        Me.dgvReclasificar = New System.Windows.Forms.DataGridView()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.cmbGrupo = New System.Windows.Forms.ComboBox()
        Me.cmbEntity = New System.Windows.Forms.ComboBox()
        Me.cmbDescrip = New System.Windows.Forms.ComboBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.btnAplicarFiltro = New System.Windows.Forms.Button()
        Me.btnQuitarFiltro = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.groupApp.SuspendLayout()
        CType(Me.dgvPolizasHFM, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvCuentaOrigen, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvReclasificar, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'groupApp
        '
        Me.groupApp.Controls.Add(Me.txtCurrDate)
        Me.groupApp.Controls.Add(Me.txtValue)
        Me.groupApp.Controls.Add(Me.txtPeriod)
        Me.groupApp.Controls.Add(Me.txtYear)
        Me.groupApp.Controls.Add(Me.txtScenario)
        Me.groupApp.Controls.Add(Me.txtHora)
        Me.groupApp.Controls.Add(Me.txtFecha)
        Me.groupApp.Controls.Add(Me.txtUsuario)
        Me.groupApp.Controls.Add(Me.Label8)
        Me.groupApp.Controls.Add(Me.Label7)
        Me.groupApp.Controls.Add(Me.Label6)
        Me.groupApp.Controls.Add(Me.Label5)
        Me.groupApp.Controls.Add(Me.Label4)
        Me.groupApp.Controls.Add(Me.Label3)
        Me.groupApp.Controls.Add(Me.Label2)
        Me.groupApp.Controls.Add(Me.Label1)
        Me.groupApp.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.groupApp.Location = New System.Drawing.Point(98, 24)
        Me.groupApp.Name = "groupApp"
        Me.groupApp.Size = New System.Drawing.Size(1006, 180)
        Me.groupApp.TabIndex = 0
        Me.groupApp.TabStop = False
        Me.groupApp.Text = "Aplicación: PRDIZZI"
        '
        'txtCurrDate
        '
        Me.txtCurrDate.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtCurrDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCurrDate.Location = New System.Drawing.Point(777, 43)
        Me.txtCurrDate.Name = "txtCurrDate"
        Me.txtCurrDate.ReadOnly = True
        Me.txtCurrDate.Size = New System.Drawing.Size(209, 27)
        Me.txtCurrDate.TabIndex = 15
        '
        'txtValue
        '
        Me.txtValue.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtValue.Location = New System.Drawing.Point(777, 131)
        Me.txtValue.Name = "txtValue"
        Me.txtValue.ReadOnly = True
        Me.txtValue.Size = New System.Drawing.Size(209, 27)
        Me.txtValue.TabIndex = 14
        '
        'txtPeriod
        '
        Me.txtPeriod.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtPeriod.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtPeriod.Location = New System.Drawing.Point(563, 131)
        Me.txtPeriod.Name = "txtPeriod"
        Me.txtPeriod.ReadOnly = True
        Me.txtPeriod.Size = New System.Drawing.Size(100, 27)
        Me.txtPeriod.TabIndex = 13
        '
        'txtYear
        '
        Me.txtYear.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtYear.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtYear.Location = New System.Drawing.Point(377, 131)
        Me.txtYear.Name = "txtYear"
        Me.txtYear.ReadOnly = True
        Me.txtYear.Size = New System.Drawing.Size(100, 27)
        Me.txtYear.TabIndex = 12
        '
        'txtScenario
        '
        Me.txtScenario.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtScenario.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtScenario.Location = New System.Drawing.Point(160, 132)
        Me.txtScenario.Name = "txtScenario"
        Me.txtScenario.ReadOnly = True
        Me.txtScenario.Size = New System.Drawing.Size(145, 27)
        Me.txtScenario.TabIndex = 11
        '
        'txtHora
        '
        Me.txtHora.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtHora.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtHora.Location = New System.Drawing.Point(160, 99)
        Me.txtHora.Name = "txtHora"
        Me.txtHora.ReadOnly = True
        Me.txtHora.Size = New System.Drawing.Size(145, 27)
        Me.txtHora.TabIndex = 10
        '
        'txtFecha
        '
        Me.txtFecha.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtFecha.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFecha.Location = New System.Drawing.Point(160, 66)
        Me.txtFecha.Name = "txtFecha"
        Me.txtFecha.ReadOnly = True
        Me.txtFecha.Size = New System.Drawing.Size(177, 27)
        Me.txtFecha.TabIndex = 9
        '
        'txtUsuario
        '
        Me.txtUsuario.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtUsuario.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtUsuario.Location = New System.Drawing.Point(160, 34)
        Me.txtUsuario.Name = "txtUsuario"
        Me.txtUsuario.ReadOnly = True
        Me.txtUsuario.Size = New System.Drawing.Size(254, 27)
        Me.txtUsuario.TabIndex = 8
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(705, 43)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(66, 20)
        Me.Label8.TabIndex = 7
        Me.Label8.Text = "Versión"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(720, 135)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(51, 20)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "Value"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(500, 133)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(57, 20)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "Period"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(328, 135)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(43, 20)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Year"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(77, 133)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(75, 20)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Scenario"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(108, 101)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(46, 20)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Hora"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(99, 73)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(55, 20)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Fecha"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(87, 41)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(67, 20)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Usuario"
        '
        'dgvPolizasHFM
        '
        Me.dgvPolizasHFM.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvPolizasHFM.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvPolizasHFM.Location = New System.Drawing.Point(95, 297)
        Me.dgvPolizasHFM.Name = "dgvPolizasHFM"
        Me.dgvPolizasHFM.RowHeadersWidth = 51
        Me.dgvPolizasHFM.RowTemplate.Height = 24
        Me.dgvPolizasHFM.Size = New System.Drawing.Size(1129, 263)
        Me.dgvPolizasHFM.TabIndex = 1
        '
        'dgvCuentaOrigen
        '
        Me.dgvCuentaOrigen.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvCuentaOrigen.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvCuentaOrigen.Location = New System.Drawing.Point(95, 597)
        Me.dgvCuentaOrigen.Name = "dgvCuentaOrigen"
        Me.dgvCuentaOrigen.RowHeadersWidth = 51
        Me.dgvCuentaOrigen.RowTemplate.Height = 24
        Me.dgvCuentaOrigen.Size = New System.Drawing.Size(1129, 84)
        Me.dgvCuentaOrigen.TabIndex = 2
        '
        'dgvReclasificar
        '
        Me.dgvReclasificar.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvReclasificar.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvReclasificar.Location = New System.Drawing.Point(95, 722)
        Me.dgvReclasificar.Name = "dgvReclasificar"
        Me.dgvReclasificar.RowHeadersWidth = 51
        Me.dgvReclasificar.RowTemplate.Height = 24
        Me.dgvReclasificar.Size = New System.Drawing.Size(1129, 148)
        Me.dgvReclasificar.TabIndex = 3
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(95, 278)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(83, 16)
        Me.Label9.TabIndex = 4
        Me.Label9.Text = "Pólizas HFM"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(95, 578)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(92, 16)
        Me.Label10.TabIndex = 5
        Me.Label10.Text = "Cuenta Origen"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(95, 703)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(78, 16)
        Me.Label11.TabIndex = 6
        Me.Label11.Text = "Reclasificar"
        '
        'cmbGrupo
        '
        Me.cmbGrupo.FormattingEnabled = True
        Me.cmbGrupo.Location = New System.Drawing.Point(71, 25)
        Me.cmbGrupo.Name = "cmbGrupo"
        Me.cmbGrupo.Size = New System.Drawing.Size(164, 24)
        Me.cmbGrupo.TabIndex = 7
        '
        'cmbEntity
        '
        Me.cmbEntity.FormattingEnabled = True
        Me.cmbEntity.Location = New System.Drawing.Point(299, 25)
        Me.cmbEntity.Name = "cmbEntity"
        Me.cmbEntity.Size = New System.Drawing.Size(115, 24)
        Me.cmbEntity.TabIndex = 8
        '
        'cmbDescrip
        '
        Me.cmbDescrip.FormattingEnabled = True
        Me.cmbDescrip.Location = New System.Drawing.Point(531, 25)
        Me.cmbDescrip.Name = "cmbDescrip"
        Me.cmbDescrip.Size = New System.Drawing.Size(333, 24)
        Me.cmbDescrip.TabIndex = 9
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(21, 28)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(44, 16)
        Me.Label12.TabIndex = 10
        Me.Label12.Text = "Grupo"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(254, 28)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(39, 16)
        Me.Label13.TabIndex = 11
        Me.Label13.Text = "Entity"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(446, 28)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(79, 16)
        Me.Label14.TabIndex = 12
        Me.Label14.Text = "Descripción"
        '
        'btnAplicarFiltro
        '
        Me.btnAplicarFiltro.Location = New System.Drawing.Point(894, 11)
        Me.btnAplicarFiltro.Name = "btnAplicarFiltro"
        Me.btnAplicarFiltro.Size = New System.Drawing.Size(92, 23)
        Me.btnAplicarFiltro.TabIndex = 13
        Me.btnAplicarFiltro.Text = "Aplicar"
        Me.btnAplicarFiltro.UseVisualStyleBackColor = True
        '
        'btnQuitarFiltro
        '
        Me.btnQuitarFiltro.Location = New System.Drawing.Point(894, 36)
        Me.btnQuitarFiltro.Name = "btnQuitarFiltro"
        Me.btnQuitarFiltro.Size = New System.Drawing.Size(92, 23)
        Me.btnQuitarFiltro.TabIndex = 14
        Me.btnQuitarFiltro.Text = "Quitar"
        Me.btnQuitarFiltro.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.btnQuitarFiltro)
        Me.GroupBox1.Controls.Add(Me.cmbGrupo)
        Me.GroupBox1.Controls.Add(Me.btnAplicarFiltro)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.cmbDescrip)
        Me.GroupBox1.Controls.Add(Me.Label14)
        Me.GroupBox1.Controls.Add(Me.cmbEntity)
        Me.GroupBox1.Location = New System.Drawing.Point(98, 210)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1006, 65)
        Me.GroupBox1.TabIndex = 15
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Filtro"
        '
        'FrmPolizasHFM
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1299, 903)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.dgvReclasificar)
        Me.Controls.Add(Me.dgvCuentaOrigen)
        Me.Controls.Add(Me.dgvPolizasHFM)
        Me.Controls.Add(Me.groupApp)
        Me.Name = "FrmPolizasHFM"
        Me.Text = "FrmPolizasHFM"
        Me.groupApp.ResumeLayout(False)
        Me.groupApp.PerformLayout()
        CType(Me.dgvPolizasHFM, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvCuentaOrigen, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvReclasificar, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents groupApp As GroupBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents txtCurrDate As TextBox
    Friend WithEvents txtValue As TextBox
    Friend WithEvents txtPeriod As TextBox
    Friend WithEvents txtYear As TextBox
    Friend WithEvents txtScenario As TextBox
    Friend WithEvents txtHora As TextBox
    Friend WithEvents txtFecha As TextBox
    Friend WithEvents txtUsuario As TextBox
    Friend WithEvents dgvPolizasHFM As DataGridView
    Friend WithEvents dgvCuentaOrigen As DataGridView
    Friend WithEvents dgvReclasificar As DataGridView
    Friend WithEvents Label9 As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents Label11 As Label
    Friend WithEvents cmbGrupo As ComboBox
    Friend WithEvents cmbEntity As ComboBox
    Friend WithEvents cmbDescrip As ComboBox
    Friend WithEvents Label12 As Label
    Friend WithEvents Label13 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents btnAplicarFiltro As Button
    Friend WithEvents btnQuitarFiltro As Button
    Friend WithEvents GroupBox1 As GroupBox
End Class
