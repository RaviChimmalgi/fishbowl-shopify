Imports System.Net
Imports System.IO
Imports System.IO.MemoryStream
Imports System.Xml

Public Class ShopifyMethods
    Public api_key As String = "870814b8a4d2972c59664f7aab7d0ab2"
    Public password As String = "12fa4b06ff34714262f7758f444cdfce"

    'Dim api_key As String = "870814b8a4d2972c59664f7aab7d0ab2"
    'Dim psword As String = "12fa4b06ff34714262f7758f444cdfce"
    'Dim SharedSecret As String = "90311acf5d7cb1b36f1b27e7d21f60be"
    'Dim storeurl As String = "@smwebserviceproject.myshopify.com/"

    Public Class shopifyorder
        Public orderid As String
        Public ordername As String
        Public createdAt As String

        'tax info
        Public taxrate As String
        Public taxname As String

        'billing info
        Public billName As String
        Public billAddress As String
        Public billCity As String
        Public billZip As String

        'shipping info
        Public shipName As String
        Public shipAddress As String
        Public shipZip As String
        Public shipCountry As String
        Public shipState As String

        Public totalPrice As String
        Public title As String
        Public fishbowlSku As String
        Public quantity As String
        ' Public itemlist As New List(Of orderitems)
    End Class

    Public Class orderitems
        Public title As String
        Public fishbowlSku As String
        Public quantity As String
        Public price As String

    End Class
    Public Class variantsToUpdate
        Public variantid As String
        Public quantity As String
    End Class


    Public Function getOrders() As String

        Dim client As New WebClient
        client.Credentials = New NetworkCredential(api_key, password)

        Dim stream As IO.Stream = client.OpenRead("https://smwebserviceproject.myshopify.com/admin/orders.xml?since_id=150583138")
        Dim reader As New StreamReader(stream)
        Dim s As String = reader.ReadToEnd


        Return s
    End Function

    Public Function getOrderDetails(ByVal xdoc As Xml.XmlDocument) As List(Of shopifyorder)
        Dim orderlist As New List(Of shopifyorder)

        Dim orderNodeList As XmlNodeList = xdoc.SelectNodes("/orders/order")

        Console.WriteLine(orderNodeList.Count)
        For Each ordernode In orderNodeList

            Dim order As New shopifyorder
            order.orderid = ordernode.Childnodes.Item(7).innertext
            order.orderName = ordernode.Childnodes.Item(29).innertext
            order.createdAt = ordernode.Childnodes.Item(20).innertext
            order.totalPrice = ordernode.Childnodes.Item(16).innertext

            Dim billNodeList As XmlNodeList = xdoc.SelectNodes("/orders/order/billing-address")
            For Each billNode In billNodeList
                order.billName = billNode.Childnodes.Item(12).innertext
                order.billAddress = billNode.Childnodes.Item(2).innertext + " " + billNode.Childnodes.Item(3).innertext + " " + billNode.Childnodes.Item(4).innertext
                order.billCity = billNode.Childnodes.Item(4).innertext
                order.billZip = billNode.Childnodes.Item(9).innertext
            Next

            Dim shipNodeList As XmlNodeList = xdoc.SelectNodes("/orders/order/shipping-address")
            For Each shipNode In shipNodeList
                order.shipName = shipNode.Childnodes.Item(12).innertext
                order.shipAddress = shipNode.Childnodes.Item(2).innertext + " " + shipNode.Childnodes.Item(3).innertext
                order.shipZip = shipNode.Childnodes.Item(9).innertext
                order.shipCountry = shipNode.Childnodes.Item(6).innertext
                order.shipState = shipNode.Childnodes.Item(8).innertext
            Next

            Dim itemNodeList As XmlNodeList = xdoc.SelectNodes("/orders/order/line-items/line-item")
            For Each itemNode In itemNodeList
                order.title = itemNode.Childnodes.Item(7).innertext
                order.fishbowlSku = itemNode.Childnodes.Item(6).innertext
                order.quantity = itemNode.Childnodes.Item(5).innertext

            Next
            Dim taxNode As XmlNode = xdoc.SelectSingleNode("/orders/order/tax-lines/tax-line")
            order.taxname = taxNode.ChildNodes.Item(0).InnerText
            order.taxrate = taxNode.ChildNodes.Item(2).InnerText
            orderlist.Add(order)

        Next


        Return orderlist
    End Function

    Public Sub updateShopifyVariants(ByVal variantlist As List(Of variantsToUpdate))

        For Each var In variantlist
            ' Create a request for the URL.  
            Dim request As WebRequest = WebRequest.Create("https://smwebserviceproject.myshopify.com/admin/variants/259224462.xml")
            request.Credentials = New NetworkCredential(api_key, password)

            Dim buffer As System.Text.StringBuilder = New System.Text.StringBuilder("")
            buffer.Append("<?xml version=""1.0"" encoding=""UTF-8""?>")
            buffer.Append("<variant>")
            buffer.Append("<id type=""integer"">" + var.variantid + "</id>")
            buffer.Append("<inventory-quantity type=""integer"">" + var.quantity + "</inventory-quantity>")
            buffer.Append("</variant>")

            Dim str As String = buffer.ToString()

            request.Method = "PUT"

            'important
            request.ContentType = "application/xml"


            ' Turn the XML into a byte buffer to prepare it for transmission 
            Dim bytebuffer As Byte() = System.Text.Encoding.UTF8.GetBytes(str)
            request.ContentLength = bytebuffer.Length
            Dim streamData As Stream = request.GetRequestStream()
            streamData.Write(bytebuffer, 0, bytebuffer.Length)
            streamData.Close()


            Dim response As HttpWebResponse = Nothing
            Try
                ' This is where the HTTP POST actually occurs.
                response = DirectCast(request.GetResponse(), HttpWebResponse)
            Catch a As Exception
                Console.WriteLine(a.ToString())
            End Try
            ' Display the status.
            Console.WriteLine(response.StatusDescription)
            ' Get the stream containing content returned by the server.
            Dim dataStream As Stream = response.GetResponseStream()
            ' Open the stream using a StreamReader for easy access.
            Dim reader As New StreamReader(dataStream)
            ' Read the content.
            Dim responseFromServer As String = reader.ReadToEnd()
            ' Display the content.
            Console.WriteLine(responseFromServer)
            ' Cleanup the streams and the response.
            reader.Close()
            dataStream.Close()
            response.Close()

        Next
    End Sub



    ' This code snippet was taken from another button I made in a Visual Studio 2010 Beta 2 Windows Forms Application.
    ' The code POSTs the XML representing a new product to a shop. After the code executes, the product should be visible via admin.
    Public Sub AddNewProductShopify()
        ' Create a request for the URL.         
        Dim request As HttpWebRequest = WebRequest.Create("https://smwebserviceproject.myshopify.com/admin/products.xml")
        request.Credentials = New NetworkCredential(api_key, password)

        ' This is the XML that represents the new product. 
        ' Note that the double quotes are a way for VS to allow quotes in a string literal
        Dim toSend As String = "<?xml version=""1.0"" encoding=""UTF-8""?>" & _
                                "<product>" & _
                                "  <product-type>Snowboard</product-type>" & _
                                "  <body>Good snowboard!</body>" & _
                                "  <title>TEMP PRODUCT</title>" & _
                                "  <variants type=""array"">" & _
                                "<variant>" & _
                                "  <option1>Second</option1>" & _
                                "  <price>20.00</price>" & _
                                "</variant>" & _
                                "  </variants>" & _
                                "  <vendor>Burton</vendor>" & _
                                "</product>"
        ' Prepare the WebRequest to POST 
        request.Method = "POST"
        ' This ended up being really important! The code won't work without it
        request.ContentType = "application/xml"

        ' Turn the XML into a byte buffer to prepare it for transmission 
        Dim lbPostBuffer As Byte() = System.Text.Encoding.UTF8.GetBytes(toSend)
        request.ContentLength = lbPostBuffer.Length
        Dim loPostData As Stream = request.GetRequestStream()
        loPostData.Write(lbPostBuffer, 0, lbPostBuffer.Length)
        loPostData.Close()

        Dim response As HttpWebResponse = Nothing
        Try
            ' This is where the HTTP POST actually occurs.
            response = DirectCast(request.GetResponse(), HttpWebResponse)
        Catch a As Exception
            Console.WriteLine(a.ToString())
        End Try
        ' Display the status.
        Console.WriteLine(response.StatusDescription)
        ' Get the stream containing content returned by the server.
        Dim dataStream As Stream = response.GetResponseStream()
        ' Open the stream using a StreamReader for easy access.
        Dim reader As New StreamReader(dataStream)
        ' Read the content.
        Dim responseFromServer As String = reader.ReadToEnd()
        ' Display the content.
        Console.WriteLine(responseFromServer)
        ' Cleanup the streams and the response.
        reader.Close()
        dataStream.Close()
        response.Close()
    End Sub
End Class