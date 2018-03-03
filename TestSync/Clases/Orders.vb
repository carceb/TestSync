Imports MySql.Data.MySqlClient
Public Class Orders
    Private claseSQL As String
    Private objetoAccessHelper As New AccessSqlHelper
    Private objetoMySqlHekper As New MySqlHelper

    Private objetoMySqlInventoryPricing As New MySqlInventoryPricing
    Private objectLibrary As New Library
    Private objetoXML As New ManejoXML
    Public Sub New()
    End Sub
    Public Sub SincronizarOrders()
        Dim readerDatos As OleDb.OleDbDataReader
        Dim readerOrderMySql As MySqlDataReader
        Dim readerOrderAccess As OleDb.OleDbDataReader
        Dim objetoInventory As New Inventory
        Dim totalRegistrosActualizados As Integer = 0
        Dim totalRegistrosNuevos As Integer = 0

        objectLibrary.WriteErrorLog("Servicio de sincronización: Sincronizando tabla = Orders ")
        objectLibrary.WriteProcessLog("Sincronizando tabla = Orders", "Orders.txt")

        objectLibrary.WriteProcessLog("Dias hacia atras " & Convert.ToInt32(objetoXML.ObtenerValorXML("CantidadDiasRestaDiaActual")), "Orders.txt")

        'Carga el readerDatos con los registros a sincronizar (ingresar o actualizar) en MySQL
        readerDatos = objetoInventory.ObtenerLastUpdate("Orders")
        '************************************************************
        Try
            If readerDatos.HasRows Then
                Do While readerDatos.Read
                    'objetoMySqlInventoryPricing.EliminaSkunoInventoryPricingMySQL(readerDatos.Item("skuno"))
                    readerOrderMySql = ObtenerMySqlOrders(readerDatos.Item("OrdID"))
                    If Not readerOrderMySql.Read Then

                    End If
                    'objetoMySqlInventoryPricing.IngresarNuevoInventoryPricing(readerInventoryPrincingAccess)
                    totalRegistrosNuevos = totalRegistrosNuevos + 1
                Loop
                readerDatos.Close()
                objetoAccessHelper.cnnOLEDB.Close()
                objectLibrary.WriteProcessLog("Inventory Pricing: Registros NUEVOS sincronizados para actualización = " & totalRegistrosNuevos, "InventoryPricing.txt")
            Else
                'objectLibrary.WriteProcessLog("Inventory Pricing: No se encontraron registros para sincronizar en la fecha = " & fechaCortaLastUpDate, "InventoryPricing.txt")
            End If
        Catch ex As Exception
            objectLibrary.WriteErrorLog(ex.Message)
        End Try
    End Sub
    Public Function ObtenerMySqlOrders(codigoOrder As String) As MySqlDataReader
        Dim otroMySqlsHelper As New MySqlHelper
        claseSQL = "SELECT * from Orders where OrdID = " & codigoOrder
        ObtenerMySqlOrders = otroMySqlsHelper.MySqlHelperExecuteReader(claseSQL)
    End Function
End Class
