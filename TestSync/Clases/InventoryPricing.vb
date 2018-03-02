Imports MySql.Data.MySqlClient
Public Class InventoryPricing 'Administra el proceso de sincronizacion de la tabla InventoryPricing en MySQL
    Private claseSQL As String
    Private objetoAccessHelper As New AccessSqlHelper
    Private objetoMySqlHekper As New MySqlHelper

    Private objetoMySqlInventoryPricing As New MySqlInventoryPricing
    Private fechaLastUpDate As DateTime
    Private fechaCortaLastUpDate As String
    Private objectLibrary As New Library
    Private objetoXML As New ManejoXML
    Public Sub New()
    End Sub
    Public Sub SincronizarInventoryPricing ()
        Dim readerDatos As OleDb.OleDbDataReader
        Dim readerInventoryPrincingAccess As OleDb.OleDbDataReader
        Dim totalRegistrosActualizados As Integer = 0
        Dim totalRegistrosNuevos As Integer = 0
        Dim diasHaciaAtras As Integer = 0

        objectLibrary.WriteErrorLog("Servicio de sincronización: Sincronizando tabla = InventoryPricing ")
        objectLibrary.WriteProcessLog("Sincronizando tabla = InventoryPricing", "InventoryPricing.txt")

        'Obtiene el numero de dias hacia atras para colocarlo como parametro del campo LastUpDate, tala Inventory de Access
        diasHaciaAtras = Convert.ToInt32(objetoXML.ObtenerValorXML("CantidadDiasRestaDiaActual"))
        fechaLastUpDate = Date.Now.AddDays(-diasHaciaAtras)
        fechaLastUpDate = fechaLastUpDate.ToShortDateString
        fechaCortaLastUpDate = fechaLastUpDate.ToString("M/d/yyyy")
        '*************************************************************
        objectLibrary.WriteProcessLog("Dias hacia atras " & diasHaciaAtras, "InventoryPricing.txt")

        'Carga el readerDatos con los registros a sincronizar (ingresar o actualizar) en MySQL
        readerDatos = ObtenerLastUpdate("Inventory")
        '************************************************************
        Try
            If readerDatos.HasRows Then
                Do While readerDatos.Read
                    objetoMySqlInventoryPricing.EliminaSkunoInvetoryPricingMySQL(readerDatos.Item("skuno"))
                    readerInventoryPrincingAccess = ObtenerAccessInventoryPricing(readerDatos.Item("skuno"))
                    objetoMySqlInventoryPricing.IngresarNuevoInventoryPricing(readerInventoryPrincingAccess)
                    totalRegistrosNuevos = totalRegistrosNuevos + 1
                Loop
                readerDatos.Close()
                objectLibrary.WriteProcessLog("Inventory Pricing: Registros NUEVOS sincronizados para actualización = " & totalRegistrosNuevos, "InventoryPricing.txt")
            Else
                objectLibrary.WriteProcessLog("Inventory Pricing: No se encontraron registros para sincronizar en la fecha = " & fechaCortaLastUpDate, "InventoryPricing.txt")
            End If
        Catch ex As Exception
            objectLibrary.WriteErrorLog(ex.Message)
        End Try
    End Sub
    Public Function ObtenerLastUpdate(nombreTabla As String) As OleDb.OleDbDataReader
        'Obtiene los registros filtrados por la cantidad de dias hacia adelante segun el parametro del campo LastUpDate 
        claseSQL = "SELECT * FROM " & nombreTabla & " Where LastUpdate   > #" & fechaCortaLastUpDate & "#"
        ObtenerLastUpdate = objetoAccessHelper.OLDBReader(claseSQL)
    End Function
    Public Function ObtenerAccessInventoryPricing(skuno As String) As OleDb.OleDbDataReader
        Dim otroAccessHelper As New AccessSqlHelper
        'Obtiene los registros filtrados por la cantidad de dias hacia adelante segun el parametro del campo LastUpDate 
        claseSQL = "Select * from [Inventory Pricing] where skuno = " & skuno
        ObtenerAccessInventoryPricing = otroAccessHelper.OLDBReader(claseSQL)
    End Function
End Class
