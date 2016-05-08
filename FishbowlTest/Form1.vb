
Imports FishbowlAPI
Imports FishbowlAPI.ShopifyMethods

Public Class Form1
    Dim X As New FishbowlAPI.FishbowlAPI(My.Settings.Username, My.Settings.Password, My.Settings.Host, My.Settings.Port)
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        ToolStripStatusLabel1.Text = "Processing"
        Dim X As New FishbowlAPI.FishbowlAPI(My.Settings.Username, My.Settings.Password, My.Settings.Host, My.Settings.Port)
        Dim DS As DataSet = X.GetParts()
        TextBox1.Text = DS.GetXml
        'DataGridView1.DataSource = DS.Tables("LightPart")
        'DS = X.GetInventory(DS.Tables("LightPart").Rows(1)("Num"))
        'DataGridView1.DataSource = DS.Tables("Location")
        'TextBox1.Text = DS.GetXml
        X.Close()
        ToolStripStatusLabel1.Text = "Done"
    End Sub

    Dim orders As New List(Of shopifyorder)
    Private Sub Get_ShopifyOrders(sender As System.Object, e As System.EventArgs) Handles GetShopifyOrders.Click
        Dim methods As New ShopifyMethods
        Dim result As String = methods.getOrders()
        TextBox1.Text = result
        Dim xdoc As New Xml.XmlDocument
        xdoc.LoadXml(result)

        orders = methods.getOrderDetails(xdoc)

        TextBox1.Text = result
    End Sub

    Private Sub Put_OrdersInFishBowl(sender As System.Object, e As System.EventArgs) Handles PutOrdersInFishBowl.Click
        Dim X As New FishbowlAPI.FishbowlAPI(My.Settings.Username, My.Settings.Password, My.Settings.Host, My.Settings.Port)
        X.AddSalesOrder(orders)
        X.Close()
    End Sub

    Private Sub Get_FishBowlOrders(sender As System.Object, e As System.EventArgs) Handles GetFishBowlOrders.Click
        Dim X As New FishbowlAPI.FishbowlAPI(My.Settings.Username, My.Settings.Password, My.Settings.Host, My.Settings.Port)
        Dim datetim As String = "1"
        Dim DS As DataSet = X.GetSalesOrder()
        TextBox1.Text = DS.GetXml
        X.Close()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Dim methods As New ShopifyMethods
        Dim var As New variantsToUpdate
        var.variantid = "259224462"
        var.quantity = "40"

        Dim varlist As New List(Of variantsToUpdate)
        varlist.Add(var)

        methods.updateShopifyVariants(varlist)
    End Sub
End Class
