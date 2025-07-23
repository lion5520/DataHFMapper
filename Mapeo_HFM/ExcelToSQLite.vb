
Imports OfficeOpenXml
Imports System.Data
Imports System.Data.SQLite
Imports System.IO
Imports System.Text
Imports System.Globalization

Module ExcelToSQLite



    Public Function LeerPrimeraHojaEPPlus(rutaExcel As String, rutaSQLite As String, rutaTxtSalida As String) As DataTable
        Dim resultado As New DataTable()
        resultado.Columns.Add("CIA", GetType(String))
        resultado.Columns.Add("S_NEG", GetType(String))
        resultado.Columns.Add("CTA", GetType(String))
        resultado.Columns.Add("DESCRIP", GetType(String))
        resultado.Columns.Add("ICP", GetType(String))
        resultado.Columns.Add("MON", GetType(String))
        resultado.Columns.Add("TIPO", GetType(String))
        resultado.Columns.Add("YEAR", GetType(Integer))
        resultado.Columns.Add("MES", GetType(String))
        resultado.Columns.Add("TOP", GetType(Integer))
        resultado.Columns.Add("IMPORTE", GetType(Decimal))

        Dim mesAnterior As String = CultureInfo.CreateSpecificCulture("es-MX").DateTimeFormat.GetAbbreviatedMonthName(DateTime.Now.AddMonths(-1).Month).Replace(".", "").ToUpper()
        Dim anioActual As Integer = DateTime.Now.Year

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial
        Using package As New ExcelPackage(New FileInfo(rutaExcel))
            Dim hoja = package.Workbook.Worksheets.First()
            Dim filaInicio As Integer = 0

            For fila As Integer = 1 To hoja.Dimension.End.Row
                If hoja.Cells(fila, 1).Text.Trim() = "Soc SIFIC" Then
                    filaInicio = fila + 1
                    Exit For
                End If
            Next

            If filaInicio = 0 Then
                Throw New Exception("No se encontraron encabezados esperados.")
            End If

            Using sw As New System.IO.StreamWriter(rutaTxtSalida, False, System.Text.Encoding.UTF8)
                sw.WriteLine("CIA" & vbTab & "S_NEG" & vbTab & "CTA" & vbTab & "DESCRIP" & vbTab & "ICP" & vbTab & "MON" & vbTab & "TIPO" & vbTab & "YEAR" & vbTab & "MES" & vbTab & "TOP" & vbTab & "IMPORTE")

                For fila As Integer = filaInicio To hoja.Dimension.End.Row
                    If hoja.Cells(fila, 1).Text.Trim().ToUpper() = "Total C14" Then Exit For

                    Dim cia = hoja.Cells(fila, 1).Text.Trim()
                    Dim cta = hoja.Cells(fila, 2).Text.Trim()
                    Dim descrip = "" 'hoja.Cells(fila, 3).Text.Trim().Trim(Chr(34))

                    Dim icp As String = hoja.Cells(fila, 3).Text.Trim()
                    Dim icpSific As String = icp

                    If icp <> "[ICP None]" Then
                        Using conexion As New SQLiteConnection("Data Source=" & rutaSQLite & ";Version=3;Read Only=False;")

                            Try
                                Using conn As New SQLiteConnection("Data Source=" & rutaSQLite & ";Version=3;")
                                    conexion.Open()
                                End Using
                            Catch ex As Exception
                                MessageBox.Show("ERROR al abrir la base de datos: " & rutaSQLite & vbCrLf & ex.Message)
                            End Try

                            Using cmd As New SQLiteCommand("SELECT Soc FROM tabla_icp_sific WHERE Original = @original", conexion)
                                cmd.Parameters.AddWithValue("@original", icp)
                                Dim icpResultado = cmd.ExecuteScalar()
                                If icpResultado IsNot Nothing Then
                                    icpSific = icpResultado.ToString().Trim()
                                End If
                            End Using
                        End Using
                    Else
                        icpSific = "000"
                    End If
                    icp = icpSific

                    Dim importeText = hoja.Cells(fila, 7).Text.Trim().Replace(",", "")
                    Dim importeDecimal As Decimal = importeText
                    'If Not Decimal.TryParse(importeText, importeDecimal) Then Continue For
                    'importeDecimal = Math.Round(importeDecimal, 2)

                    'If String.IsNullOrWhiteSpace(cta) And String.IsNullOrWhiteSpace(descrip) Then Continue For

                    Dim top As Integer = 0
                    If icp.ToUpper() <> "[ICP NONE]" And icp.ToUpper() <> "000" Then
                        Using conn As New SQLiteConnection("Data Source=" & rutaSQLite & ";Version=3;")
                            conn.Open()
                            Dim cmd As New SQLiteCommand("SELECT TOP_UNICO, ESTADO FROM variacion_top WHERE CIA = @cia AND CTA = @cta AND ICIA = @icp AND ESTADO = 'Sin cambio'", conn)
                            cmd.Parameters.AddWithValue("@cia", cia)
                            cmd.Parameters.AddWithValue("@cta", cta)
                            cmd.Parameters.AddWithValue("@icp", icp)

                            Dim TOP_UNICO = "0"
                            Dim ESTADO = "0"
                            Using reader = cmd.ExecuteReader()
                                While reader.Read()
                                    TOP_UNICO = reader("TOP_UNICO").ToString().Trim()
                                    ESTADO = reader("ESTADO").ToString().Trim()

                                End While
                            End Using

                            If ESTADO <> "0" Then
                                If ESTADO = "Sin cambio" Then
                                    top = Convert.ToInt32(TOP_UNICO)
                                ElseIf ESTADO = "Con variación" Then
                                    top = -2
                                End If

                            Else
                                top = -3
                            End If

                            'Dim resultadoTop = "1" 'cmd.ExecuteScalar()
                            'Dim resultadoEstado = cmd.ExecuteReader(2)

                            'If resultadoTop IsNot Nothing Then

                            '    top = Convert.ToInt32(resultadoTop)
                            'Else
                            '    top = -3
                            'End If
                        End Using
                    End If

                    resultado.Rows.Add(cia, " ", cta, descrip, icp, "MXN", "H", anioActual, mesAnterior, top, importeDecimal)
                    sw.WriteLine(String.Join(vbTab, {
                    cia,
                    "2000",
                    cta,
                    descrip,
                    icp,
                    "MXN",
                    "H",
                    anioActual.ToString(),
                    mesAnterior,
                    top.ToString(),
                    Math.Round(importeDecimal, 2).ToString("F2", System.Globalization.CultureInfo.InvariantCulture)
                }))
                Next
            End Using
        End Using

        Return resultado
    End Function




End Module
