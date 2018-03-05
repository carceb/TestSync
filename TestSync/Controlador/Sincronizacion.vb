
Public Class Sincronizacion 'Controla todos los objetos encargados de la sincronizacion en la base de datos
    Public Sub IniciarProcesoSincronizacion()
        Dim objetoInventory As New Inventory
        Dim objetoInventoryPricing As New InventoryPricing
        Dim objetoInventoryItemXRef As New InventoryItemXRef
        Dim objetoOrders As New Orders
        Dim objetoOrdersDetail As New OrdersDetail
        Dim objetoInventoryDTS As New InventoryDTS
        Dim objetoCustomersBill As New CustomersBill
        Dim objetoCustomersBillPrsn As New CustomersBillPrsn
        Dim objectLibrary As New Library

        'objetoInventory.SincronizarInventory()
        'objetoInventoryPricing.SincronizarInventoryPricing()
        'objetoInventoryItemXRef.SincronizarInventoryItemXRef()
        'objetoOrders.SincronizarOrders()
        'objetoOrdersDetail.SincronizarOrdersDatail()
        'objetoCustomersBill.SincronizarCustomersBill()

        objetoCustomersBillPrsn.SincronizarCustomersBillPrsn()

        'Los campos que estan el codigo guia no son los de la tabla se debe investigar como es este proceso
        'objetoInventoryDTS.SincronizarInventoryDTS()

    End Sub
End Class

