
Public Class Sincronizacion 'Controla todos los objetos encargados de la sincronizacion en la base de datos
    Private objetoInventory As New Inventory
    Private objetoInventoryPricing As New InventoryPricing
    Private objetoInventoryItemXRef As New InventoryItemXRef
    Private objectLibrary As New Library
    Public Sub IniciarProcesoSincronizacion()
        'objetoInventory.SincronizarInventory()
        objetoInventoryPricing.SincronizarInventoryPricing()
        ' objetoInventoryItemXRef.SincronizarInventoryItemXRef()
    End Sub
End Class

