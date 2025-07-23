Imports System.Data
Imports System.Data.SQLite
Imports System.Drawing

Public Class CompletaDatosHFM

    Private ReadOnly _rutaSqlite As String
    Private _dtDatos As DataTable

    Public Sub New(rutaSqlite As String)
        InitializeComponent()
        _rutaSqlite = rutaSqlite
        AddHandler Me.Load, AddressOf CompletaDatosHFM_Load
        AddHandler Me.FormClosing, AddressOf CompletaDatosHFM_FormClosing
    End Sub

    ''' <summary>
    ''' Al cargar, cierra si no hay faltantes, o carga la grilla.
    ''' </summary>
    Private Sub CompletaDatosHFM_Load(sender As Object, e As EventArgs)
        If ContarFaltantes() = 0 Then
            Me.DialogResult = DialogResult.OK
            Me.Close()
        Else
            CargarDatos()
        End If
    End Sub

    ''' <summary>
    ''' Cuenta cuántos registros siguen sin HFM u Oracle.
    ''' </summary>
    Private Function ContarFaltantes() As Integer
        Using conn As New SQLiteConnection($"Data Source={_rutaSqlite};Version=3;")
            conn.Open()
            Using cmd As New SQLiteCommand(
                "SELECT COUNT(1) FROM t_in_sap
                 WHERE TRIM(cuenta_mayor_hfm) = ''
                    OR TRIM(cuenta_oracle) = ''", conn)
                Return Convert.ToInt32(cmd.ExecuteScalar())
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Carga sólo los registros pendientes y oculta el ID.
    ''' </summary>
    Private Sub CargarDatos()
        Dim sql = "SELECT id, sociedad, numero_cuenta, texto_explicativo, deudor_acreedor, saldo_acum, " &
                  "cuenta_mayor_hfm, descripcion_cuenta_sific, cuenta_oracle, descripcion_cuenta_oracle " &
                  "FROM t_in_sap " &
                  "WHERE TRIM(cuenta_mayor_hfm) = '' OR TRIM(cuenta_oracle) = ''"

        Using conn As New SQLiteConnection($"Data Source={_rutaSqlite};Version=3;")
            conn.Open()
            Dim da As New SQLiteDataAdapter(sql, conn)
            _dtDatos = New DataTable()
            da.Fill(_dtDatos)
        End Using

        dgvDatos.DataSource = _dtDatos

        For Each col As DataGridViewColumn In dgvDatos.Columns
            Select Case col.Name.ToLower()
                Case "id"
                    col.Visible = False
                Case "sociedad", "numero_cuenta", "texto_explicativo", "deudor_acreedor", "saldo_acum"
                    col.ReadOnly = True : col.Visible = True
                Case "cuenta_mayor_hfm", "descripcion_cuenta_sific", "cuenta_oracle", "descripcion_cuenta_oracle"
                    col.ReadOnly = False : col.Visible = True
                    col.DefaultCellStyle.BackColor = Color.White
                Case Else
                    col.Visible = False
            End Select
        Next
    End Sub

    ''' <summary>
    ''' Al cambiar una celda editable, pinta la fila.
    ''' </summary>
    Private Sub dgvDatos_CellValueChanged(s As Object, e As DataGridViewCellEventArgs) _
        Handles dgvDatos.CellValueChanged
        If e.RowIndex < 0 Then Return
        dgvDatos.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.LightYellow
    End Sub

    ''' <summary>
    ''' Guarda los cambios y, si aún hay faltantes, recarga la grilla.
    ''' </summary>
    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Dim filasMod As DataRow() = _dtDatos.Select(Nothing, Nothing, DataViewRowState.ModifiedCurrent)
        If filasMod.Length = 0 Then
            MessageBox.Show("No hay cambios para guardar.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Using conn As New SQLiteConnection($"Data Source={_rutaSqlite};Version=3;")
            conn.Open()
            Using tx = conn.BeginTransaction()
                Dim cmd As New SQLiteCommand(
                    "UPDATE t_in_sap SET 
                        cuenta_mayor_hfm = @cmhfm, 
                        descripcion_cuenta_sific = @descSific, 
                        cuenta_oracle = @cOracle, 
                        descripcion_cuenta_oracle = @descOracle 
                     WHERE id = @id", conn, tx)

                cmd.Parameters.Add(New SQLiteParameter("@cmhfm", DbType.String))
                cmd.Parameters.Add(New SQLiteParameter("@descSific", DbType.String))
                cmd.Parameters.Add(New SQLiteParameter("@cOracle", DbType.String))
                cmd.Parameters.Add(New SQLiteParameter("@descOracle", DbType.String))
                cmd.Parameters.Add(New SQLiteParameter("@id", DbType.Int32))

                For Each row As DataRow In filasMod
                    cmd.Parameters("@cmhfm").Value = row("cuenta_mayor_hfm").ToString()
                    cmd.Parameters("@descSific").Value = row("descripcion_cuenta_sific").ToString()
                    cmd.Parameters("@cOracle").Value = row("cuenta_oracle").ToString()
                    cmd.Parameters("@descOracle").Value = row("descripcion_cuenta_oracle").ToString()
                    cmd.Parameters("@id").Value = CInt(row("id"))
                    cmd.ExecuteNonQuery()
                Next

                tx.Commit()
            End Using
        End Using

        ' 🔄 Volvemos a contar y recargar o cerrar
        Dim faltantes = ContarFaltantes()
        If faltantes > 0 Then
            MessageBox.Show(
                $"Quedan {faltantes} registro(s) por completar.",
                "Atención",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning)
            CargarDatos()
        Else
            MessageBox.Show(
                "Todos los registros han sido completados.",
                "Éxito",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End If
    End Sub

    ''' <summary>
    ''' Cerrar = cancelar => devuelve Cancel
    ''' </summary>
    Private Sub btnCancelar_Click(sender As Object, e As EventArgs) Handles btnCancelar.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    ''' <summary>
    ''' Si cierra por X, forzamos DialogResult.Cancel.
    ''' </summary>
    Private Sub CompletaDatosHFM_FormClosing(sender As Object, e As FormClosingEventArgs)
        If Me.DialogResult <> DialogResult.OK Then
            Me.DialogResult = DialogResult.Cancel
        End If
    End Sub

End Class
