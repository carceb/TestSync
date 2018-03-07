Public Class Form1
    Private objetoAccessHelper As New AccessSqlHelper
    Private objetoXML As New ManejoXML
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Dim sql As String
        'Dim x As New MySqlHelper
        'sql = "INSERT INTO factura(FacturaID,ClienteID,MaterialID,MontoFactura) VALUES(11,2,2,5000)"
        'x.create(sql)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim objetoSincronizacion As New Sincronizacion
        Dim intervaloTimer As Double
        Dim objetoConf As New ManejoXML
        Dim numeroCompilacion As Integer
        intervaloTimer = TiempoEjecucionSincronizador(objetoXML.ObtenerValorXML("TiempoEjecucionServicioSincronizacion", "Configuracion.xml"))

        'Validar cantidad de ejecuciones
        '****************************************************************************************************
        If objetoConf.ActualizarSBDPXML() = True Then
            numeroCompilacion = Convert.ToInt32(objetoConf.ObtenerValorXML("Compilaciones", "SBDP.xml"))
            If numeroCompilacion >= 480 Or numeroCompilacion = 0 Then
                MessageBox.Show("Error en el modulo .NET 4.5")
                End
            End If
        End If
        '*********************************************************************************************

        objetoSincronizacion.IniciarProcesoSincronizacion()
        MessageBox.Show("Fin")
    End Sub
End Class
