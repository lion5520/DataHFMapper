Imports System.IO
Imports System.Data.SQLite
Imports Microsoft.Office.Interop.Excel

Public Class ReporteICLoader

    Private ReadOnly _rutaBd As String
    Private ReadOnly _rutaExcel As String

    ''' <summary>
    ''' Inicializa el loader con la ruta de la BD SQLite y la ruta del archivo Excel a procesar.
    ''' </summary>
    ''' <param name="rutaBd">Ruta al archivo .sqlite.</param>
    ''' <param name="rutaExcel">Ruta al archivo .xlsx que contiene la pestaña de saldos.</param>
    Public Sub New(rutaBd As String, rutaExcel As String)
        If String.IsNullOrWhiteSpace(rutaBd) Then
            Throw New ArgumentException("La ruta de la BD no puede estar vacía.", NameOf(rutaBd))
        End If
        If String.IsNullOrWhiteSpace(rutaExcel) OrElse Not File.Exists(rutaExcel) Then
            Throw New ArgumentException("La ruta del Excel no es válida o el archivo no existe.", NameOf(rutaExcel))
        End If

        _rutaBd = rutaBd
        _rutaExcel = rutaExcel
    End Sub

    ''' <summary>
    ''' Crea la tabla (si no existe), limpia los datos existentes (y reinicia la secuencia AUTOINCREMENT),
    ''' y carga los registros de la hoja "saldos" incluyendo el nuevo campo ICSap.
    ''' </summary>
    Public Sub Cargar()
        Using conn As New SQLiteConnection($"Data Source={_rutaBd}")
            conn.Open()

            '' 1) Crear tabla si no existe
            'Using cmd As New SQLiteCommand(conn)
            '    cmd.CommandText =
            '        "CREATE TABLE IF NOT EXISTS reporte_IC (" &
            '        "   id INTEGER PRIMARY KEY AUTOINCREMENT, " &
            '        "   ICSap TEXT NOT NULL, " &
            '        "   SociedadSap TEXT NOT NULL, " &
            '        "   CuentaSap TEXT NOT NULL, " &
            '        "   Saldo REAL" &
            '        ");"
            '    cmd.ExecuteNonQuery()
            'End Using

            ' 2) Borrar datos existentes
            Using cmdClear As New SQLiteCommand("DELETE FROM reporte_IC;", conn)
                cmdClear.ExecuteNonQuery()
            End Using

            ' 3) Reiniciar AUTOINCREMENT
            Using cmdSeq As New SQLiteCommand(
                    "DELETE FROM sqlite_sequence WHERE name='reporte_IC';", conn)
                cmdSeq.ExecuteNonQuery()
            End Using

            ' 4) Abrir Excel y ubicar hoja "saldos"
            Dim xlApp As New Application
            Dim wb As Workbook = xlApp.Workbooks.Open(_rutaExcel, ReadOnly:=True)
            Dim ws As Worksheet = wb.Sheets("saldos")

            ' 5) Detectar fila de encabezados y columnas de interés
            Dim headerRow As Integer = -1
            Dim colICSap As Integer = -1
            Dim colSociedad As Integer = -1
            Dim colCuenta As Integer = -1
            Dim colCuentaOracle As Integer = -1
            Dim colSaldo As Integer = -1

            For r As Integer = 1 To 10
                For c As Integer = 1 To ws.UsedRange.Columns.Count
                    Dim txt = ws.Cells(r, c).Value2
                    If txt IsNot Nothing Then
                        Select Case txt.ToString().Trim().ToUpperInvariant()
                            Case "IC SAP" : colICSap = c
                            Case "SOCIEDAD SAP" : colSociedad = c
                            Case "CUENTA SAP" : colCuenta = c
                            Case "CUENTA ORACLE" : colCuentaOracle = c
                            Case "SALDO" : colSaldo = c
                        End Select
                    End If
                Next
                If colICSap > 0 AndAlso colSociedad > 0 _
                   AndAlso colCuenta > 0 AndAlso colSaldo > 0 Then
                    headerRow = r
                    Exit For
                End If
            Next

            If headerRow < 0 Then
                wb.Close(False) : xlApp.Quit()
                Throw New InvalidOperationException(
                    "No se encontró la fila de encabezados con 'IC SAP', " &
                    "'SOCIEDAD SAP', 'CUENTA SAP', 'CUENTA ORACLE' y 'SALDO'.")
            End If

            ' 6) Recorrer filas posteriores hasta bloque vacío
            Dim fila As Integer = headerRow + 1
            While True
                Dim vIC = ws.Cells(fila, colICSap).Value2
                Dim vSoc = ws.Cells(fila, colSociedad).Value2
                Dim vCta = ws.Cells(fila, colCuenta).Value2
                Dim vCtaOra = ws.Cells(fila, colCuentaOracle).Value2
                Dim vSld = ws.Cells(fila, colSaldo).Value2

                ' Si todas son Nothing, fin de datos
                If vIC Is Nothing AndAlso vSoc Is Nothing _
                   AndAlso vCta Is Nothing AndAlso vSld Is Nothing Then
                    Exit While
                End If

                ' Transformar y validar
                Dim icsap As String = If(vIC?.ToString().Trim(), "")
                Dim sociedad As String = If(vSoc?.ToString().Trim(), "")
                Dim cuenta As String = If(vCta?.ToString().Trim(), "")
                Dim cuentaOracle As String = If(vCtaOra?.ToString().Trim(), "")
                Dim saldo As Double
                If vSld IsNot Nothing AndAlso Not Double.TryParse(vSld.ToString(), saldo) Then
                    saldo = 0
                End If

                ' Insertar si sociedad y cuenta no están vacíos
                If sociedad <> "" AndAlso cuenta <> "" Then
                    Using cmdIns As New SQLiteCommand(conn)
                        cmdIns.CommandText =
                            "INSERT INTO reporte_IC (ICSap, SociedadSap, CuentaSap, Cuenta_Parte_Relacionada, Saldo) " &
                            "VALUES (@ic, @soc, @cta, @vCtaOra, @sld);"
                        cmdIns.Parameters.AddWithValue("@ic", icsap)
                        cmdIns.Parameters.AddWithValue("@soc", sociedad)
                        cmdIns.Parameters.AddWithValue("@cta", cuenta)
                        cmdIns.Parameters.AddWithValue("@vCtaOra", cuentaOracle)
                        cmdIns.Parameters.AddWithValue("@sld", saldo)
                        cmdIns.ExecuteNonQuery()
                    End Using
                End If

                fila += 1
            End While

            ' 7) Cerrar Excel
            wb.Close(False)
            xlApp.Quit()
        End Using
    End Sub

End Class
