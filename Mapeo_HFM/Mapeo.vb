
Imports System.Data
Imports System.Data.SQLite
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Text
Imports OfficeOpenXml

Public Class Mapeo

    Private rutaSQLite_A As String = "C:\Users\Lion\Desktop\IZZI_DAT\mapeo_sap_sific.sqlite"
    Private rutaSQLite As String = "Y:\top_variacion.sqlite"
    Private originalImage_SAP_IN As Image
    Private originalImage_procesa_1 As Image
    Private originalImage_previsualiza_1 As Image
    Private originalImage_previsualiza_2 As Image
    Private originalImage_previsualiza_3 As Image
    Private originalImage_previsualiza_4 As Image
    Private originalImage_previsualiza_5 As Image
    Private originalImage_flecha_1 As Image
    Private originalImage_flecha_2 As Image
    Private originalImage_flecha_2_1 As Image
    Private originalImage_flecha_2_2 As Image
    Private originalImage_flecha_3 As Image
    Private originalImage_procesa_2 As Image
    Private originalImage_txt_sific As Image

    Private Sub guardaImagenesOrignals()
        originalImage_SAP_IN = SAP_IN.Image
        originalImage_procesa_1 = procesa_1.Image
        originalImage_previsualiza_1 = previsualiza_1.Image
        originalImage_previsualiza_2 = previsualiza_2.Image
        originalImage_previsualiza_3 = previsualiza_3.Image
        originalImage_previsualiza_4 = previsualiza_4.Image
        originalImage_previsualiza_5 = previsualiza_5.Image
        originalImage_flecha_1 = flecha_1.Image
        originalImage_flecha_2 = flecha_2.Image
        originalImage_flecha_2_1 = flecha_2_1.Image
        originalImage_flecha_2_2 = flecha_2_2.Image
        originalImage_flecha_3 = flecha_3.Image
        originalImage_procesa_2 = procesa_2.Image
        originalImage_txt_sific = txt_sific.Image
    End Sub

    Private Sub todoGris()
        SAP_IN.Image = ToGrayscale(originalImage_SAP_IN)
        procesa_1.Image = ToGrayscale(originalImage_procesa_1)
        previsualiza_1.Image = ToGrayscale(originalImage_previsualiza_1)
        previsualiza_2.Image = ToGrayscale(originalImage_previsualiza_2)
        previsualiza_3.Image = ToGrayscale(originalImage_previsualiza_3)
        previsualiza_4.Image = ToGrayscale(originalImage_previsualiza_4)
        previsualiza_5.Image = ToGrayscale(originalImage_previsualiza_5)
        flecha_1.Image = ToGrayscale(originalImage_flecha_1)
        flecha_2.Image = ToGrayscale(originalImage_flecha_2)
        flecha_2_1.Image = ToGrayscale(originalImage_flecha_2_1)
        flecha_2_2.Image = ToGrayscale(originalImage_flecha_2_2)
        flecha_3.Image = ToGrayscale(originalImage_flecha_3)
        procesa_2.Image = ToGrayscale(originalImage_procesa_2)
        txt_sific.Image = ToGrayscale(originalImage_txt_sific)
    End Sub

    Private Sub desavilitaTodo()
        procesa_1.Enabled = False
        previsualiza_1.Enabled = False
        previsualiza_2.Enabled = False
        previsualiza_3.Enabled = False
        previsualiza_4.Enabled = False
        previsualiza_5.Enabled = False
        flecha_1.Enabled = False
        flecha_2.Enabled = False
        flecha_2_1.Enabled = False
        flecha_2_2.Enabled = False
        flecha_3.Enabled = False
        procesa_2.Enabled = False
        txt_sific.Enabled = False
    End Sub
    Private Sub Mapeo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Guarda la imagen original para poder restaurarla
        guardaImagenesOrignals()
        'Convierte todo en gris nada ejecutado
        todoGris()

    End Sub

    Private Sub btnCargaFile_Click(sender As Object, e As EventArgs) Handles btnCargaFile.Click
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Archivos Excel|*.xlsx;*.xls"
        openFileDialog.Title = "Selecciona el archivo Excel"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim rutaExcel As String = openFileDialog.FileName

            ' Primero preguntamos la ruta de guardado
            Dim saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "Archivos de texto|*.txt"
            saveFileDialog.Title = "Guardar archivo de salida"

            IniciarParpadeoBoton(btnCargaFile)

            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                Dim rutaSalida As String = saveFileDialog.FileName

                ' Ejecutamos el procesamiento y generación del archivo
                Dim tablaDatos As DataTable = LeerPrimeraHojaEPPlus(rutaExcel, rutaSQLite, rutaSalida)

                If tablaDatos.Rows.Count > 0 Then
                    MessageBox.Show("Archivo TXT guardado exitosamente.")
                Else
                    MessageBox.Show("No se encontraron datos para exportar.")
                End If
            End If
        End If

        DetenerParpadeo()

    End Sub




    Private ParpadeoCancelToken As Threading.CancellationTokenSource

    Private Async Sub IniciarParpadeoBoton(boton As Button)
        ParpadeoCancelToken = New Threading.CancellationTokenSource()
        Dim token = ParpadeoCancelToken.Token

        Dim colorOriginal = boton.BackColor
        Dim colorAlterno = Color.Orange

        Try
            While Not token.IsCancellationRequested
                boton.BackColor = If(boton.BackColor = colorOriginal, colorAlterno, colorOriginal)
                Await Task.Delay(500, token) ' Cambia cada 0.5 segundos
            End While
        Catch ex As TaskCanceledException
            boton.BackColor = colorOriginal ' Restaurar color al cancelar
        End Try
    End Sub

    Private Sub DetenerParpadeo()
        If ParpadeoCancelToken IsNot Nothing Then
            ParpadeoCancelToken.Cancel()
            ParpadeoCancelToken = Nothing
        End If
    End Sub

    Private Async Sub IniciarParpadeoImagen(PicImagen As PictureBox)
        ' No arranques dos veces
        If ParpadeoCancelToken IsNot Nothing Then Return

        PicImagen.Enabled = True

        ParpadeoCancelToken = New Threading.CancellationTokenSource()
        Dim token = ParpadeoCancelToken.Token

        ' Guarda originales
        Dim originalImg = PicImagen.Image
        Dim grayImg = ToGrayscale(originalImg) ' tu función ya existente

        Try
            Dim mostrarOriginal As Boolean = False

            While Not token.IsCancellationRequested
                ' Alterna entre original y gris
                PicImagen.Image = If(mostrarOriginal, originalImg, grayImg)
                mostrarOriginal = Not mostrarOriginal
                PicImagen.Refresh()
                Await Task.Delay(500, token)
            End While

        Catch ex As TaskCanceledException
            ' cancelado: nada que hacer
        Finally
            ' Restaura la original
            PicImagen.Image = originalImg
        End Try

    End Sub


    Private Async Sub SAP_IN_Click(sender As Object, e As EventArgs) Handles SAP_IN.Click
        Dim rutaExcel As String



        ' 1) Selección del archivo Excel
        Using dlg As New OpenFileDialog() With {
            .Title = "Seleccione el archivo Excel",
            .Filter = "Archivos Excel|*.xlsx;*.xls"
        }
            If dlg.ShowDialog() <> DialogResult.OK Then Exit Sub
            rutaExcel = dlg.FileName
        End Using



        '' 2) Selección de la base SQLite
        'Using dlg As New OpenFileDialog() With {
        '    .Title = "Seleccione la base SQLite",
        '    .Filter = "SQLite Database|*.db;*.sqlite"
        '}
        '    If dlg.ShowDialog() <> DialogResult.OK Then Exit Sub
        '    rutaSQLite_A = dlg.FileName
        'End Using


        ' 1) Arranca el parpadeo
        SAP_IN.Image = originalImage_SAP_IN
        IniciarParpadeoImagen(SAP_IN)

        Try
            Me.Cursor = Cursors.WaitCursor
            ' 2) Ejecuta tu operación en segundo plano
            Await Task.Run(Sub()
                               Procesa_Entrada_SAP.Ejecutar(rutaExcel, rutaSQLite_A)
                               MessageBox.Show("Proceso finalizado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)

                               'Inicia rutina si hay cuenta oracle vacias

                               ' 2) Validas y, si hay datos faltantes, muestras el formulario
                               Using frm As New CompletaDatosHFM(rutaSQLite_A)
                                   If frm.ShowDialog() = DialogResult.OK Then
                                       ' Si hubo faltantes, el usuario los completó; 
                                       ' si no había nada, el formulario ni siquiera apareció.
                                   End If
                               End Using


                           End Sub)

        Catch ex As Exception
            MessageBox.Show("Ocurrió un error:" & Environment.NewLine & ex.Message,
                                                              "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' 3) Detiene el parpadeo (restaura la imagen original)
            DetenerParpadeo()
            Me.Cursor = Cursors.Default
        End Try


        'Si todo bien pinta en color 
        flecha_1.Image = originalImage_flecha_1

        procesa_1.Cursor = Cursors.Hand
        procesa_1.Enabled = True

    End Sub



    ''' <summary>
    ''' Devuelve un Bitmap convertido a escala de grises.
    ''' </summary>
    Private Function ToGrayscale(ByVal src As Image) As Bitmap
        Dim bmp = New Bitmap(src.Width, src.Height)

        ' Matriz para transformar a escala de grises
        Dim cm As New ColorMatrix(New Single()() {
            New Single() {0.3F, 0.3F, 0.3F, 0, 0},
            New Single() {0.59F, 0.59F, 0.59F, 0, 0},
            New Single() {0.11F, 0.11F, 0.11F, 0, 0},
            New Single() {0, 0, 0, 1, 0},
            New Single() {0, 0, 0, 0, 1}
        })

        Using g As Graphics = Graphics.FromImage(bmp)
            Dim ia As New ImageAttributes()
            ia.SetColorMatrix(cm)
            ' Dibuja la imagen original aplicando la matriz de color
            g.DrawImage(src,
                        New Rectangle(0, 0, bmp.Width, bmp.Height),
                        0, 0, src.Width, src.Height,
                        GraphicsUnit.Pixel,
                        ia)
        End Using

        Return bmp
    End Function

    Private Sub previsualiza_1_Click(sender As Object, e As EventArgs) Handles previsualiza_1.Click

        Dim exporter = New SqliteTableExporter(rutaSQLite_A, "t_in_sap", "Previsualiza_5")
        exporter.Export()
    End Sub

    Private Sub b_salir_Click(sender As Object, e As EventArgs) Handles b_salir.Click
        End
    End Sub

    Private Async Sub procesa_1_Click(sender As Object, e As EventArgs) Handles procesa_1.Click
        If procesa_1.Enabled = True Then

            ' 2) Validas y, si hay datos faltantes, muestras el formulario
            'Using frm As New CompletaDatosHFM(rutaSQLite_A)
            '    Dim result = frm.ShowDialog()
            '    If result = DialogResult.OK Then
            '        ' El usuario completó todo: continúa...
            '    Else
            '        ' El usuario cerró/canceló antes de terminar: aborta tu flujo
            '        Return
            '    End If
            'End Using


            'Procesa_CIA_ICP.Ejecutar(rutaSQLite_A)



            ' 1) Arranca el parpadeo
            procesa_1.Image = originalImage_procesa_1
            IniciarParpadeoImagen(procesa_1)

            Try
                Me.Cursor = Cursors.WaitCursor
                ' 2) Ejecuta tu operación en segundo plano
                Await Task.Run(Sub()
                                   Dim proc As New AperturaDetalleProcessor(rutaSQLite_A)
                                   proc.ProcesarReporteIC()

                               End Sub)

            Catch ex As Exception
                MessageBox.Show("Ocurrió un error:" & Environment.NewLine & ex.Message,
                                                              "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                ' 3) Detiene el parpadeo (restaura la imagen original)
                DetenerParpadeo()
                Me.Cursor = Cursors.Default
            End Try




            'Si todo bien pinta en color 
            procesa_1.Image = originalImage_procesa_1
            flecha_2.Image = originalImage_flecha_2
            flecha_2_1.Image = originalImage_flecha_2_1
            flecha_2_2.Image = originalImage_flecha_2_2
            previsualiza_1.Image = originalImage_previsualiza_1

            procesa_2.Cursor = Cursors.Hand
            procesa_2.Enabled = True

            previsualiza_1.Enabled = True
            previsualiza_1.Cursor = Cursors.Hand

            lbl_conversion_icp.ForeColor = Color.ForestGreen
            lbl_conversion_cia.ForeColor = Color.ForestGreen
            lbl_conversion_cia.Visible = True
            lbl_conversion_icp.Visible = True

        End If
    End Sub

    Private Sub previsualiza_General_Click(sender As Object, e As EventArgs) Handles previsualiza_General.Click
        Using dlg As New MultiLineInputForm()
            dlg.TextValue = "SELECT 
sociedad,
SUM(saldo_acum) AS total_saldo_acum
FROM t_in_sap
WHERE saldo_acum IS NOT NULL
GROUP BY sociedad
ORDER BY sociedad;
"   ' valor por defecto
            If dlg.ShowDialog(Me) <> DialogResult.OK Then
                Return   ' Canceló
            End If

            Dim sql = dlg.TextValue.Trim()
            If String.IsNullOrWhiteSpace(sql) Then
                MessageBox.Show("La consulta no puede estar vacía.", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Try
                ExcelDbExporter.ExportToExcel(rutaSQLite_A, fullQuery:=sql)
            Catch ex As Exception
                MessageBox.Show($"Ocurrió un error:{vbCrLf}{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub



    ' Formulario para texto multilínea
    Public Class MultiLineInputForm
        Inherits Form

        Public Property TextValue As String
            Get
                Return txtInput.Text
            End Get
            Set(value As String)
                txtInput.Text = value
            End Set
        End Property

        Private txtInput As New TextBox() With {
        .Multiline = True,
        .ScrollBars = ScrollBars.Vertical,
        .Dock = DockStyle.Fill,
        .Font = New Font("Consolas", 10),
        .AcceptsReturn = True
    }
        Private btnAceptar As New Button() With {
        .Text = "Aceptar",
        .DialogResult = DialogResult.OK,
        .Dock = DockStyle.Bottom,
        .Height = 30
    }
        Private btnCancelar As New Button() With {
        .Text = "Cancelar",
        .DialogResult = DialogResult.Cancel,
        .Dock = DockStyle.Bottom,
        .Height = 30
    }

        Public Sub New()
            Me.Text = "Introducir consulta SQL"
            Me.StartPosition = FormStartPosition.CenterParent
            Me.Width = 600
            Me.Height = 300
            Me.Controls.Add(txtInput)
            Me.Controls.Add(btnAceptar)
            Me.Controls.Add(btnCancelar)
            Me.AcceptButton = btnAceptar
            Me.CancelButton = btnCancelar
        End Sub
    End Class

    Private Async Sub procesa_2_Click(sender As Object, e As EventArgs) Handles procesa_2.Click
        ' 1) Arranca el parpadeo
        procesa_2.Image = originalImage_procesa_2
        IniciarParpadeoImagen(procesa_2)

        Try
            Me.Cursor = Cursors.WaitCursor
            ' 2) Ejecuta tu operación en segundo plano
            Await Task.Run(Sub()
                               ' 2) Crea la instancia
                               Dim agrupador As New agrupa_cuenta_mayor(rutaSQLite_A)
                               ' 3) Llama al proceso
                               agrupador.Procesar()
                               MessageBox.Show("Proceso finalizado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                           End Sub)

        Catch ex As Exception
            MessageBox.Show("Ocurrió un error:" & Environment.NewLine & ex.Message,
                                                              "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' 3) Detiene el parpadeo (restaura la imagen original)
            DetenerParpadeo()
            Me.Cursor = Cursors.Default
        End Try


        'Si todo bien pinta en color 
        flecha_3.Image = originalImage_flecha_3

        procesa_2.Cursor = Cursors.Hand
        procesa_2.Enabled = True
        previsualiza_4.Enabled = True
        previsualiza_4.Image = originalImage_previsualiza_4

        txt_sific.Enabled = True
        txt_sific.Cursor = Cursors.Hand


    End Sub

    Private Sub previsualiza_4_Click(sender As Object, e As EventArgs) Handles previsualiza_4.Click

        Dim exporter = New SqliteTableExporter(rutaSQLite_A, "t_in_sap", "Previsualiza_5")
        exporter.Export()
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click

        Dim frm As New Captura_Polizas(rutaSQLite_A, Me)
        frm.Show()

    End Sub

    Private Async Sub previsualiza_2_Click(sender As Object, e As EventArgs) Handles previsualiza_2.Click

        Dim rutaExcel As String



        ' 1) Selección del archivo Excel
        Using dlg As New OpenFileDialog() With {
            .Title = "Seleccione el archivo Excel",
            .Filter = "Archivos Excel|*.xlsx;*.xls"
        }
            If dlg.ShowDialog() <> DialogResult.OK Then Exit Sub
            rutaExcel = dlg.FileName
        End Using


        ' 1) Arranca el parpadeo
        previsualiza_2.Image = originalImage_previsualiza_2
        IniciarParpadeoImagen(previsualiza_2)

        Try
            Me.Cursor = Cursors.WaitCursor
            ' 2) Ejecuta tu operación en segundo plano
            Await Task.Run(Sub()

                               'Inicia rutina de carga complemento operaciones IC
                               Dim loader As New ReporteICLoader(rutaSQLite_A, rutaExcel)
                               loader.Cargar()
                               MessageBox.Show("Proceso finalizado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)


                           End Sub)

        Catch ex As Exception
            MessageBox.Show("Ocurrió un error:" & Environment.NewLine & ex.Message,
                                                              "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' 3) Detiene el parpadeo (restaura la imagen original)
            DetenerParpadeo()
            Me.Cursor = Cursors.Default
        End Try


        'Si todo bien pinta en color 
        flecha_2_1.Image = originalImage_flecha_2_1

        txt_sific.Cursor = Cursors.Hand
        txt_sific.Enabled = True






    End Sub

    Private Async Sub txt_sific_Click(sender As Object, e As EventArgs) Handles txt_sific.Click


        ' 1) Arranca el parpadeo
        txt_sific.Image = originalImage_txt_sific
        IniciarParpadeoImagen(txt_sific)


        Try
            Me.Cursor = Cursors.WaitCursor
            ' 2) Ejecuta tu operación en segundo plano
            Await Task.Run(Sub()

                               'Inicia rutina de carga complemento operaciones IC
                               Dim transformer As New LayOutTransformerICB(rutaSQLite_A)

                               transformer.Transform()
                               MessageBox.Show("Datos transformados OK.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)


                           End Sub)


        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' 3) Detiene el parpadeo (restaura la imagen original)
            DetenerParpadeo()
            Me.Cursor = Cursors.Default
        End Try

        'Si todo bien pinta en color 
        txt_sific.Image = originalImage_txt_sific

        previsualiza_5.Enabled = True
        previsualiza_5.Image = originalImage_previsualiza_5

    End Sub

    Private Sub previsualiza_5_Click(sender As Object, e As EventArgs) Handles previsualiza_5.Click

        Dim exporter = New SqliteTableExporter(rutaSQLite_A, "t_in_sap", "Previsualiza_5")
        exporter.Export()

    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click


        Dim exporter = New SqliteTableExporter(rutaSQLite_A, "t_in_sap", "Previsualiza_5")
        exporter.Export()
    End Sub

End Class
