Imports System.Security.Cryptography
Imports System.Xml
Imports System.IO
Imports FishbowlAPI.ShopifyMethods
Imports System.Data.SqlClient

Public Class FishbowlAPI
    Dim Conn1 As New SqlConnection("Data Source=CMIE-SI06;Initial Catalog=shopify_Fishbowl;Integrated Security =SSPI")
    'Private properties
    Private ticket As String            'the login key needed to authenticate to Fishbowl server (set in constructor)
    Private ConnObj As ConnectionObject 'the tcp connection to Fishbowl server (set in constructor)

    'Constructor for class.
    Public Sub New(LoginName As String, Password As String, Host As String, Port As Integer)
        ' Construct XML login string
        Dim loginCommand As String = createLoginXml(LoginName, Password)
        ConnObj = New ConnectionObject(Host, Port)

        'Send Login Command once to get fishbowl server to recognize the connection attempt
        'or pull the key off the line if already connected
        Dim S As String = ConnObj.sendCommand(loginCommand)
        Dim key As String = pullKey(S)
        If (key = "null") Then
            S = ConnObj.sendCommand(loginCommand)
            key = pullKey(S)
        End If
        'Set ticket property
        ticket = "<Ticket><Key>" & key & "</Key></Ticket>"
    End Sub

    'GetInventory.
    'Returns all details on specific part, including QtyOnhand, QtyAvailable, QtyCommitted.
    'Note each part may have multiple locations (with separate quantities for each).
    'Quantity details are in InvQty datatable, with related child row in Location table (note is 1:1 relation though -
    'each InvQty record has 1 child Location record)
    Public Function GetInventory(PartNum As String) As DataSet
        Dim S As String = "<FbiXml>" & _
                            ticket & _
                            "<FbiMsgsRq><InvQtyRq>" & _
                            "<PartNum>" & PartNum & "</PartNum>" & _
                            "<LastModifiedFrom></LastModifiedFrom>" & _
                            "<LastModifiedTo></LastModifiedTo>" & _
                            "</InvQtyRq></FbiMsgsRq>" & _
                          "</FbiXml>"
        Dim DS As New DataSet
        DS.ReadXml(XmlReader.Create(New StringReader(IssueCommand(S))))
        Return DS
    End Function

    'GetParts.
    'Returns record for each part in inventory (in "LightPart" datatable)
    Public Function GetParts() As DataSet
        Dim S As String = "<FbiXml>" & _
                    ticket & _
                    "<FbiMsgsRq><LightPartListRq></LightPartListRq></FbiMsgsRq>" & _
                  "</FbiXml>"

        Dim DS As New DataSet
        DS.ReadXml(XmlReader.Create(New StringReader(IssueCommand(S))))
        Return DS
    End Function
    ' "<DateCreatedBegin>" + datestrin + "</DateCreatedBegin>" & _
    '                    "<DateCreatedEnd>" + datestop + "<DateCreatedEnd>" & _
    ' "<DateCreatedBegin>2012-11-26T00:00:00</DateCreatedBegin>" & _
    'getOrders
    'can differentiate between shopify and highland if set salesman to 'shopify' and 'highland' to respective orders 
    Public Function GetSalesOrder() As DataSet
        '"<DateIssuedBegin>" + datetim + "</DateIssuedBegin>"
        Dim datestrin As String = "2012-10-07T16:13:31"
        Dim datestop As String = "2012-12-10T16:13:31"
        Dim S As String = "<FbiXml>" & ticket & _
                          "<FbiMsgsRq><GetSOListRq>" & _
                         "<DateCreatedBegin>2012-12-09T00:00:00</DateCreatedBegin>" & _
                          "</GetSOListRq></FbiMsgsRq></FbiXml>"
        Dim DS As New DataSet
        DS.ReadXml(XmlReader.Create(New StringReader(IssueCommand(S))))
        Return DS

    End Function

    Public Class skuvariantid
        Public sku As String
        Public variantid As String
    End Class

    'Public Function getSkuVariantid(ByVal ds As DataSet) As List(Of skuvariantid)
    '    Dim idList As List(Of skuvariantid)
    '    Dim doc As New XmlDocument
    '    doc.LoadXml(ds.GetXml)
    'End Function

    'Status ID: 
    '10 - Entered, 
    '20 - Picking 
    '30 - Partial 
    '40 - Picked 
    '50 - Fulfilled 
    '60 – Closed Short 
    '70 - Void 
    '45 - Shipped

    'AddSalesOrder.
    'Creates a new sales order
    Public Sub AddSalesOrder(ByVal order As List(Of shopifyorder))
        '***implement
        For Each ord In order
            'Dim ord As shopifyorder = order(0)
            Dim buffer As System.Text.StringBuilder = New System.Text.StringBuilder("")
            buffer.Append("<FbiXml>" + ticket)
            buffer.Append("<FbiMsgsRq> " & _
                          "<AddSOItemRq>" & _
                          "<SalesOrder> " & _
                          "<Salesman></Salesman>")
            buffer.Append("<Number></Number>" & _
                          "<Status>10</Status>" & _
                          "<Carrier>FEDEX</Carrier>" & _
                          "<FirstShipDate></FirstShipDate>")

            buffer.Append("<CreatedDate>" + ord.createdAt + "</CreatedDate>")
            buffer.Append("<IssuedDate></IssuedDate>" & _
                          "<TaxRatePercentage>" + ord.taxrate + "</TaxRatePercentage>" & _
                          "<TaxRateName>" + ord.taxname + "</TaxRateName>" & _
                          "<ShippingTerms></ShippingTerms>" & _
                          "<PaymentTerms>COD</PaymentTerms>" & _
                          "<CustomerContact></CustomerContact>" & _
                          "<CustomerName>SouthernMarshCustomer</CustomerName>" & _
                          "<FOB></FOB>" & _
                          "<QuickBooksClassName></QuickBooksClassName>" & _
                          "<LocationGroup></LocationGroup>")

            buffer.Append("<BillTo>" & _
                          "<Name>" + ord.billName + "</Name>" & _
                          "<AddressField>" + ord.billAddress + "</AddressField>" & _
                          "<City>" + ord.billCity + "</City>" & _
                          "<Zip>" + ord.billZip + "</Zip>" & _
                          "</BillTo>")

            buffer.Append("<Ship>" & _
                          "<Name>" + ord.shipName + "</Name>" & _
                          "<AddressField>" + ord.shipAddress + "</AddressField>" & _
                          "<Zip>" + ord.shipZip + "</Zip>" & _
                          "<Country>" + ord.shipCountry + "</Country>" & _
                          "<State>" + ord.shipState + "</State>" & _
                          "</Ship>" & _
                          "</SalesOrder>")

            'loop for each salesorderitem
            buffer.Append("<SalesOrderItem>" & _
                          "<ID>1</ID>" & _
                          "<ProductNumber>" + ord.fishbowlSku + "</ProductNumber>" & _
                          "<SOID>2</SOID>" & _
                          "<Description>" + ord.title + "</Description>" & _
                          "<Taxable>true</Taxable>" & _
                          "<Quantity>" + ord.quantity + "</Quantity>" & _
                          "<ProductPrice>" + ord.totalPrice + "</ProductPrice>" & _
                          "<TotalPrice>" + ord.totalPrice + "</TotalPrice>" & _
                          "<UOMCode>ea</UOMCode>" & _
                          "<ItemType>10</ItemType>" & _
                          "<Status>10</Status>" & _
                          "<QuickBooksClassName></QuickBooksClassName>" & _
                          "<NewItemFlag>false</NewItemFlag>" & _
                          "</SalesOrderItem>")

            buffer.Append("</AddSOItemRq >" & _
                          "</FbiMsgsRq>" & _
                          "</FbiXml>")

            Console.WriteLine("Buffer: " + buffer.ToString)

            Dim DS As New DataSet
            DS.ReadXml(XmlReader.Create(New StringReader(IssueCommand(buffer.ToString))))


        Next


    End Sub

    '******Lower level functions******

    'Close.
    'Close TCP connection to Fishbowl server
    Public Sub Close()
        ConnObj.Close()
    End Sub

    'IssueCommand.
    'Issue XML command (in form of string) to Fishbow Server. XML string response is returned
    Public Function IssueCommand(Command As String) As String
        Return ConnObj.sendCommand(Command)
    End Function

    'createLoginXML.
    'Constructs XML login command as string
    Private Function createLoginXml(username As String, password As String) As String
        'Encrypt password
        Dim M As MD5 = MD5CryptoServiceProvider.Create()
        Dim encoded As Byte() = M.ComputeHash(System.Text.Encoding.ASCII.GetBytes(password))
        Dim encrypted As String = Convert.ToBase64String(encoded, 0, 16)
        'Create command string
        Dim buffer As System.Text.StringBuilder = New System.Text.StringBuilder("")
        buffer.Append("<FbiXml><Ticket/><FbiMsgsRq><LoginRq>")
        buffer.Append("<IAID>222</IAID>")
        buffer.Append("<IAName>API Call</IAName>")
        buffer.Append("<IADescription>API Call</IADescription>")
        buffer.Append("<UserName>" & username & "</UserName>")
        buffer.Append("<UserPassword>" & encrypted & "</UserPassword>")
        buffer.Append("</LoginRq></FbiMsgsRq></FbiXml>")
        Return buffer.ToString()
    End Function

    'pullKey.
    'Extracts the login key (which must be included with subsequence commands) from
    'the XML login return string
    Private Function pullKey(connection As String) As String 'Pull the session Key out of the server response string
        Dim key As String = ""
        Using reader As XmlReader = XmlReader.Create(New StringReader(connection))
            While (reader.Read())
                If (reader.Name = "Key") Then
                    reader.Read()
                    Return reader.Value.ToString()
                End If
            End While
            Return key
        End Using
    End Function

End Class
