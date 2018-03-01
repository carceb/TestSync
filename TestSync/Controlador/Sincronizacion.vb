
Public Class Sincronizacion 'Controla todos los objetos encargados de la sincronizacion en la base de datos
    Private objetoInventory As New Inventory
    Private objectLibrary As New Library
    Public Sub IniciarProcesoSincronizacion()
        objectLibrary.WriteErrorLog("Servicio de sincronización: Sincronizando tabla = Inventory")

        objetoInventory.SincronizarInventory()
    End Sub
End Class

