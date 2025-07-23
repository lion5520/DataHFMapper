Public NotInheritable Class ProgresoGlobal

    Private Shared _instancia As IProgreso

    Public Shared Sub Inicializar(instancia As IProgreso)
        _instancia = instancia
    End Sub

    Public Shared Sub Reportar(avance As Integer, mensaje As String)
        _instancia?.ActualizarProgreso(avance, mensaje)
        _instancia?.AgregarMensajeDebug(mensaje)
    End Sub

    Public Shared Sub Finalizar()
        _instancia?.Finalizar()
    End Sub

    Public Shared Sub Debug(mensaje As String)
        _instancia?.AgregarMensajeDebug(mensaje)
    End Sub


End Class
