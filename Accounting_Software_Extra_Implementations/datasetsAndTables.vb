Imports System.IO
Imports System.Data.SqlClient
Imports System.Data.OracleClient

Public Class datasetsAndTables

    ''' <summary>
    ''' FINALIZED CHECKJING
    ''' </summary>
    ''' <param name="orderIDParam"></param>
    ''' <returns></returns>
    Public Function finalizedCheck(ByVal orderIDParam As Integer) As DataSet
        Dim invoicesDataSet = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("SELECT DISTINCT
                                    invoices.invoiceNumber,
                                    invoices.deliveryMerged_ID,
                                    invoices.delivery_ID,
                                    invoices.invoiceType_ID,
                                    exported,
                                    finalized

                                    FROM easyenquiry.invoices

								    LEFT JOIN easyenquiry.delivery ON invoices.delivery_ID = delivery.deliveryID
                                    LEFT JOIN easyenquiry.orders ON delivery.order_ID = orders.orderID
                                    LEFT JOIN easyenquiry.quotations ON orders.quotation_ID = quotations.quotationID
                                    LEFT JOIN easyenquiry.enquiries ON quotations.enquiry_ID = enquiries.enquiryID
                                    LEFT JOIN easyenquiry.deliverymerged ON invoices.deliveryMerged_ID = deliverymerged.deliveryMergedID
                                    LEFT JOIN easyenquiry.deliveryitems ON deliverymerged.deliveryMergedID = deliveryitems.deliveryMerged_ID
                                    LEFT JOIN easyenquiry.quotitemslist ON deliveryitems.quotItemList_ID = quotitemslist.quotItemListID
                                    LEFT JOIN easyenquiry.suppliersinvoices ON delivery.deliveryID = suppliersinvoices.delivery_ID
                                    LEFT JOIN easyenquiry.supplierquotations ON enquiries.enquiryID = supplierquotations.enquiry_ID

                                    WHERE orderID = '" & orderIDParam & "' AND finalized = 1
                                    ORDER BY deliveryMerged_ID ASC", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(invoicesDataSet, "invoicesSelling")

        Return invoicesDataSet
    End Function

    ''' <summary>
    ''' EXPORTED CHECKING
    ''' </summary>
    ''' <param name="orderIDParam"></param>
    ''' <returns></returns>
    Public Function exportedCheck(ByVal orderIDParam As Integer) As DataSet
        Dim invoicesDataSet = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("SELECT DISTINCT
                                    invoices.invoiceNumber,
                                    invoices.deliveryMerged_ID,
                                    invoices.delivery_ID,
                                    invoices.invoiceType_ID,
                                    exported,
                                    finalized

                                    FROM easyenquiry.invoices

								    LEFT JOIN easyenquiry.delivery ON invoices.delivery_ID = delivery.deliveryID
                                    LEFT JOIN easyenquiry.orders ON delivery.order_ID = orders.orderID
                                    LEFT JOIN easyenquiry.quotations ON orders.quotation_ID = quotations.quotationID
                                    LEFT JOIN easyenquiry.enquiries ON quotations.enquiry_ID = enquiries.enquiryID
                                    LEFT JOIN easyenquiry.deliverymerged ON invoices.deliveryMerged_ID = deliverymerged.deliveryMergedID
                                    LEFT JOIN easyenquiry.deliveryitems ON deliverymerged.deliveryMergedID = deliveryitems.deliveryMerged_ID
                                    LEFT JOIN easyenquiry.quotitemslist ON deliveryitems.quotItemList_ID = quotitemslist.quotItemListID
                                    LEFT JOIN easyenquiry.suppliersinvoices ON delivery.deliveryID = suppliersinvoices.delivery_ID
                                    LEFT JOIN easyenquiry.supplierquotations ON enquiries.enquiryID = supplierquotations.enquiry_ID

                                    WHERE orderID = '" & orderIDParam & "' AND exported = 0 AND finalized = 1
                                    ORDER BY deliveryMerged_ID ASC", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(invoicesDataSet, "invoicesSelling")

        Return invoicesDataSet
    End Function

    ''' <summary>
    ''' CREDIT NOTE EXPORTED CHECKJING
    ''' </summary>
    ''' <param name="orderIDParam"></param>
    ''' <returns></returns>
    Public Function cnExportedCheck(ByVal orderIDParam As Integer) As DataSet
        Dim invoicesDataSet = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("  SELECT order_ID,
		                                    deliveryID,
		                                    creditNoteID,
		                                    creditnotes.exported

		                                    FROM easyenquiry.creditnotes

		                                    LEFT JOIN easyenquiry.delivery ON creditnotes.delivery_ID = delivery.deliveryID

		                                    WHERE order_ID = '" & orderIDParam & "' AND exported = 0", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(invoicesDataSet, "invoicesSelling")

        Return invoicesDataSet
    End Function

    ''' <summary>
    ''' CREDIT NOTE EXPORTED CHECKJING
    ''' </summary>
    ''' <param name="orderIDParam"></param>
    ''' <returns></returns>
    Public Function dnExportedCheck(ByVal orderIDParam As Integer) As DataSet
        Dim invoicesDataSet = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("  SELECT order_ID,
		                                    deliveryID,
		                                    debitNoteID,
		                                    debitnotes.exported

		                                    FROM easyenquiry.debitnotes

		                                    LEFT JOIN easyenquiry.delivery ON debitnotes.delivery_ID = delivery.deliveryID

		                                    WHERE order_ID = '" & orderIDParam & "' AND exported = 0", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(invoicesDataSet, "invoicesSelling")

        Return invoicesDataSet
    End Function

    Public Function anExportedCheck(ByVal orderIDParam As Integer) As DataSet
        Dim invoicesDataSet = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("SELECT order_ID,
		                                    deliveryID,
		                                    adjustmentID,
		                                    adjustments.exported

		                                    FROM easyenquiry.adjustments

		                                    LEFT JOIN easyenquiry.delivery ON adjustments.delivery_ID = delivery.deliveryID

		                                    WHERE order_ID = '" & orderIDParam & "' AND exported = 0", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(invoicesDataSet, "invoicesSelling")

        Return invoicesDataSet
    End Function

    ''' <summary>
    ''' Get distinct values of supplier ID and supplierinvoicenumber FOR BUYING 
    ''' </summary>
    ''' <param name="orderIDParam"></param>
    ''' <returns></returns>
    Public Function loadBuyingInvocies(ByVal orderIDParam As Integer) As DataSet
        Dim invoicesDataSet = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("SELECT *
                                                    FROM    (				SELECT Distinct
													InvoiceNumbers,
													ROW_NUMBER() OVER (PARTITION BY InvoiceNumbers, deliveryMergedID ORDER BY InvoiceNumbers, deliveryMergedID) AS RowNumber ,
													Names,
													deliveryID,
	                                                deliveryMergedID,
                                                    finalized,
                                                    exported,
													supplierCode
	                                                
                                                FROM   
                                                   (SELECT Distinct
													
                                                    deliveryID,
	                                                deliveryMergedID,
                                                    finalized,
                                                    exported,
													supplierCode,
	                                                suppliersinvoices.supplierInvoiceNumber, 
													(CASE
													WHEN (suppliersinvoices.generalSupplier1InvoiceNumber = '????INV') THEN CONCAT('????INV-1',deliveryID) ELSE suppliersinvoices.generalSupplier1InvoiceNumber END) as generalSupplier1InvoiceNumber,	 
													(CASE
													WHEN (suppliersinvoices.generalSupplier2InvoiceNumber = '????INV') THEN CONCAT('????INV-2',deliveryID) ELSE suppliersinvoices.generalSupplier2InvoiceNumber END) as
	                                                generalSupplier2InvoiceNumber, 
													(CASE
													WHEN (suppliersinvoices.generalSupplier3InvoiceNumber = '????INV') THEN CONCAT('????INV-3',deliveryID) ELSE suppliersinvoices.generalSupplier3InvoiceNumber END) as
	                                                generalSupplier3InvoiceNumber, 
													(CASE
													WHEN (suppliersinvoices.forwarderInvoiceNumber = '????INV') THEN CONCAT('????INV-4',deliveryID) ELSE suppliersinvoices.forwarderInvoiceNumber END) as
	                                                forwarderInvoiceNumber, 
													(CASE
													WHEN (suppliersinvoices.clearingAgentInvoiceNumber = '????INV') THEN CONCAT('????INV-5',deliveryID) ELSE suppliersinvoices.clearingAgentInvoiceNumber END) as
	                                                clearingAgentInvoiceNumber  
												
                                                FROM easyenquiry.easyenquiry.suppliersinvoices 
                                                LEFT JOIN easyenquiry.easyenquiry.delivery ON suppliersinvoices.delivery_ID = delivery.deliveryID
                                                LEFT JOIN easyenquiry.easyenquiry.deliverymerged ON delivery.deliveryID = deliverymerged.delivery_ID
                                                LEFT JOIN easyenquiry.easyenquiry.invoices ON deliverymerged.deliveryMergedID = invoices.deliveryMerged_ID
												LEFT JOIN easyenquiry.easyenquiry.supplierquotations ON delivery.supQuotation_ID = supplierquotations.supQuotationID
												LEFT JOIN easyenquiry.easyenquiry.suppliers ON supplierquotations.supplier_ID = suppliers.supplierID
                                                WHERE order_ID= '" & orderIDParam & "'
												) p 
                                                UNPIVOT  
                                                   (InvoiceNumbers 
	                                                FOR Names
	                                                IN   
                                                      (supplierInvoiceNumber,
	                                                  generalSupplier1InvoiceNumber, 
	                                                  generalSupplier2InvoiceNumber, 
	                                                  generalSupplier3InvoiceNumber, 
	                                                  forwarderInvoiceNumber, 
	                                                  clearingAgentInvoiceNumber)  
	                                                )AS 
	                                                unpvt 
                                                WHERE 
	                                                InvoiceNumbers <> '') AS a WHERE a.RowNumber = 1

                                                ORDER BY deliveryID ASC, Names DESC", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(invoicesDataSet, "invoicesBuying")
        Return invoicesDataSet
    End Function

    ''' <summary>
    ''' Get distinct values of supplier ID and supplierinvoicenumber FOR BUYING 
    ''' </summary>
    ''' <param name="orderIDParam"></param>
    ''' <returns></returns>
    Public Function loadBuyingInvociesNew(ByVal orderIDParam As Integer) As DataSet
        Dim invoicesDataSet = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("SELECT *
                                                    FROM    (				SELECT Distinct
													InvoiceNumbers,
													ROW_NUMBER() OVER (PARTITION BY InvoiceNumbers, deliveryMergedID ORDER BY InvoiceNumbers, deliveryMergedID) AS RowNumber ,
													Names,
													deliveryID,
	                                                deliveryMergedID,
                                                    finalized,
                                                    exported,
													supplierCode,
	                                                forwarderCode,
													clearingAgentCode,
													generalSupplier1Code,
													generalSupplier2Code,
													generalSupplier3Code,
													forwarderInvoiceDate,
                                                    dueDate2,
                                                    deliveryMethod_ID, 
                                                    supplierInvoiceDate,
                                                    generalSupplier1InvoiceDate,
                                                    supplierVat_ID

                                                FROM   
                                                   (SELECT Distinct
													
                                                    deliveryID,
	                                                deliveryMergedID,
                                                    finalized,
                                                    exported,
													supplierCode,
													(Select forwarderCode From easyenquiry.easyenquiry.forwarders Where delivery.forwarder_ID = forwarders.forwarderID) as forwarderCode,
													(Select clearingAgentCode From easyenquiry.easyenquiry.clearingagents Where delivery.clearingAgent_ID = clearingagents.clearingAgentID) as clearingAgentCode,
													(Select generalSupplierCode From easyenquiry.easyenquiry.generalsuppliers Where delivery.generalSupplier1_ID = generalsuppliers.generalSupplierID) as generalSupplier1Code,
													(Select generalSupplierCode From easyenquiry.easyenquiry.generalsuppliers Where delivery.generalSupplier2_ID = generalsuppliers.generalSupplierID) as generalSupplier2Code,
													(Select generalSupplierCode From easyenquiry.easyenquiry.generalsuppliers Where delivery.generalSupplier3_ID = generalsuppliers.generalSupplierID) as generalSupplier3Code,
	                                                suppliersinvoices.supplierInvoiceNumber, 
													suppliersinvoices.forwarderInvoiceNumber, 
													suppliersinvoices.clearingAgentInvoiceNumber,
                                                    suppliersinvoices.generalSupplier1InvoiceNumber,	 
                                                    suppliersinvoices.generalSupplier2InvoiceNumber, 
													suppliersinvoices.generalSupplier3InvoiceNumber,
													suppliersinvoices.forwarderInvoiceDate,
                                                    suppliersinvoices.supplierInvoiceDate,
                                                    suppliersinvoices.generalSupplier1InvoiceDate,
                                                    delivery.deliveryMethod_ID,
													CONVERT(VARCHAR(10),CAST(
													case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.freightPaymentTerm_ID) = 'invoiceDate' 
													then dispatchDate 
													when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =delivery.freightPaymentTerm_ID) = 'dispatchDate' 
													then suppliersinvoices.forwarderInvoiceDate else deliveryDate end  AS DATE),103) AS dispatchDate, 
													case when ((SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.freightPaymentTerm_ID) > 0) 
													then 
													case when convert(date, CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.freightPaymentTerm_ID) , 
													case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =delivery.freightPaymentTerm_ID) = 'invoiceDate' 
													then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =delivery.freightPaymentTerm_ID) = 'dispatchDate' 
													then suppliersinvoices.forwarderInvoiceDate else deliveryDate end )) AS DATE),103),103)
													< 
													convert(date, CONVERT(VARCHAR(10),CAST(suppliersinvoices.forwarderInvoiceDate AS DATE),103),103) 
													then CONVERT(VARCHAR(10),CAST(suppliersinvoices.forwarderInvoiceDate AS DATE),103) 
													else CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.freightPaymentTerm_ID) , 
													case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.freightPaymentTerm_ID) = 'invoiceDate' 
													then dispatchDate 
													when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.freightPaymentTerm_ID) = 'dispatchDate' 
													then suppliersinvoices.forwarderInvoiceDate  else deliveryDate end)) AS DATE),103) end else CONVERT(VARCHAR(10),CAST(suppliersinvoices.forwarderInvoiceDate  AS DATE),103) end AS dueDate2,
                                                    delivery.supplierVat_ID


                                                FROM easyenquiry.easyenquiry.suppliersinvoices 
                                                LEFT JOIN easyenquiry.easyenquiry.delivery ON suppliersinvoices.delivery_ID = delivery.deliveryID
                                                LEFT JOIN easyenquiry.easyenquiry.deliverymerged ON delivery.deliveryID = deliverymerged.delivery_ID
                                                LEFT JOIN easyenquiry.easyenquiry.invoices ON deliverymerged.deliveryMergedID = invoices.deliveryMerged_ID
												LEFT JOIN easyenquiry.easyenquiry.supplierquotations ON delivery.supQuotation_ID = supplierquotations.supQuotationID
												LEFT JOIN easyenquiry.easyenquiry.suppliers ON supplierquotations.supplier_ID = suppliers.supplierID
												LEFT JOIN easyenquiry.easyenquiry.orders ON delivery.order_ID = orders.orderID
												LEFT JOIN easyenquiry.easyenquiry.quotations ON orders.quotation_ID = quotations.quotationID

                                                WHERE order_ID= '" & orderIDParam & "' AND (
                                                    suppliersinvoices.supplierInvoiceNumber <> '' OR
													suppliersinvoices.forwarderInvoiceNumber <> '' OR
													suppliersinvoices.clearingAgentInvoiceNumber <> '' OR
                                                    suppliersinvoices.generalSupplier1InvoiceNumber <> '' OR	 
                                                    suppliersinvoices.generalSupplier2InvoiceNumber <> '' OR 
													suppliersinvoices.generalSupplier3InvoiceNumber <> '')
												) p 
                                                UNPIVOT  
                                                   (InvoiceNumbers 
	                                                FOR Names
	                                                IN   
                                                      (supplierInvoiceNumber,
													  forwarderInvoiceNumber, 
	                                                  clearingAgentInvoiceNumber,
	                                                  generalSupplier1InvoiceNumber, 
	                                                  generalSupplier2InvoiceNumber, 
	                                                  generalSupplier3InvoiceNumber)
	                                                )AS 
	                                                unpvt 
                                                WHERE 
	                                                InvoiceNumbers <> '') AS a WHERE a.RowNumber = 1 AND deliveryMergedID is not null

                                                ORDER BY deliveryID ASC, Names DESC", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(invoicesDataSet, "invoicesBuying")

        If invoicesDataSet.Tables(0).Rows.Count = 0 Then
            Dim cnDelivery1 As SqlConnection = New SqlConnection(SQL_CONNECT)
            Dim daFeatures1 As New SqlDataAdapter("SELECT *
                                                    FROM    (				SELECT Distinct
													InvoiceNumbers,
													ROW_NUMBER() OVER (PARTITION BY InvoiceNumbers, deliveryMergedID ORDER BY InvoiceNumbers, deliveryMergedID) AS RowNumber ,
													Names,
													deliveryID,
	                                                deliveryMergedID,
                                                    finalized,
                                                    exported,
													supplierCode,
	                                                forwarderCode,
													clearingAgentCode,
													generalSupplier1Code,
													generalSupplier2Code,
													generalSupplier3Code,
													forwarderInvoiceDate,
                                                    dueDate2,
                                                    deliveryMethod_ID, 
                                                    supplierInvoiceDate,
                                                    generalSupplier1InvoiceDate

                                                FROM   
                                                   (SELECT Distinct
													
                                                    deliveryID,
	                                                deliveryMergedID,
                                                    finalized,
                                                    exported,
													supplierCode,
													(Select forwarderCode From easyenquiry.easyenquiry.forwarders Where delivery.forwarder_ID = forwarders.forwarderID) as forwarderCode,
													(Select clearingAgentCode From easyenquiry.easyenquiry.clearingagents Where delivery.clearingAgent_ID = clearingagents.clearingAgentID) as clearingAgentCode,
													(Select generalSupplierCode From easyenquiry.easyenquiry.generalsuppliers Where delivery.generalSupplier1_ID = generalsuppliers.generalSupplierID) as generalSupplier1Code,
													(Select generalSupplierCode From easyenquiry.easyenquiry.generalsuppliers Where delivery.generalSupplier2_ID = generalsuppliers.generalSupplierID) as generalSupplier2Code,
													(Select generalSupplierCode From easyenquiry.easyenquiry.generalsuppliers Where delivery.generalSupplier3_ID = generalsuppliers.generalSupplierID) as generalSupplier3Code,
	                                                suppliersinvoices.supplierInvoiceNumber, 
													suppliersinvoices.forwarderInvoiceNumber, 
													suppliersinvoices.clearingAgentInvoiceNumber,
                                                    suppliersinvoices.generalSupplier1InvoiceNumber,	 
                                                    suppliersinvoices.generalSupplier2InvoiceNumber, 
													suppliersinvoices.generalSupplier3InvoiceNumber,
													suppliersinvoices.forwarderInvoiceDate,
                                                    suppliersinvoices.supplierInvoiceDate,
                                                    suppliersinvoices.generalSupplier1InvoiceDate,
                                                    delivery.deliveryMethod_ID,
													CONVERT(VARCHAR(10),CAST(
													case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.freightPaymentTerm_ID) = 'invoiceDate' 
													then dispatchDate 
													when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =delivery.freightPaymentTerm_ID) = 'dispatchDate' 
													then suppliersinvoices.forwarderInvoiceDate else deliveryDate end  AS DATE),103) AS dispatchDate, 
													case when ((SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.freightPaymentTerm_ID) > 0) 
													then 
													case when convert(date, CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.freightPaymentTerm_ID) , 
													case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =delivery.freightPaymentTerm_ID) = 'invoiceDate' 
													then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =delivery.freightPaymentTerm_ID) = 'dispatchDate' 
													then suppliersinvoices.forwarderInvoiceDate else deliveryDate end )) AS DATE),103),103)
													< 
													convert(date, CONVERT(VARCHAR(10),CAST(suppliersinvoices.forwarderInvoiceDate AS DATE),103),103) 
													then CONVERT(VARCHAR(10),CAST(suppliersinvoices.forwarderInvoiceDate AS DATE),103) 
													else CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.freightPaymentTerm_ID) , 
													case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.freightPaymentTerm_ID) = 'invoiceDate' 
													then dispatchDate 
													when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.freightPaymentTerm_ID) = 'dispatchDate' 
													then suppliersinvoices.forwarderInvoiceDate  else deliveryDate end)) AS DATE),103) end else CONVERT(VARCHAR(10),CAST(suppliersinvoices.forwarderInvoiceDate  AS DATE),103) end AS dueDate2


                                                FROM easyenquiry.easyenquiry.suppliersinvoices 
                                                LEFT JOIN easyenquiry.easyenquiry.delivery ON suppliersinvoices.delivery_ID = delivery.deliveryID
                                                LEFT JOIN easyenquiry.easyenquiry.deliverymerged ON delivery.deliveryID = deliverymerged.delivery_ID
                                                LEFT JOIN easyenquiry.easyenquiry.invoices ON deliverymerged.deliveryMergedID = invoices.deliveryMerged_ID
												LEFT JOIN easyenquiry.easyenquiry.supplierquotations ON delivery.supQuotation_ID = supplierquotations.supQuotationID
												LEFT JOIN easyenquiry.easyenquiry.suppliers ON supplierquotations.supplier_ID = suppliers.supplierID
												LEFT JOIN easyenquiry.easyenquiry.orders ON delivery.order_ID = orders.orderID
												LEFT JOIN easyenquiry.easyenquiry.quotations ON orders.quotation_ID = quotations.quotationID

                                                WHERE order_ID= '" & orderIDParam & "'
												) p 
                                                UNPIVOT  
                                                   (InvoiceNumbers 
	                                                FOR Names
	                                                IN   
                                                      (supplierInvoiceNumber,
													  forwarderInvoiceNumber, 
	                                                  clearingAgentInvoiceNumber,
	                                                  generalSupplier1InvoiceNumber, 
	                                                  generalSupplier2InvoiceNumber, 
	                                                  generalSupplier3InvoiceNumber)
	                                                )AS 
	                                                unpvt 
                                                WHERE 
	                                                InvoiceNumbers <> '') AS a WHERE a.RowNumber = 1 AND deliveryMergedID is not null

                                                ORDER BY deliveryID ASC, Names DESC", cnDelivery)

            daFeatures1.SelectCommand.CommandTimeout = 0

            ' data adapter filling dataset
            daFeatures1.Fill(invoicesDataSet, "invoicesBuying")
        End If

        Return invoicesDataSet
    End Function

    ''' <summary>
    ''' Get distinct values of supplier ID and supplierinvoicenumber FOR BUYING ************** NOT USED *********************
    ''' </summary>
    ''' <param name="orderIDParam"></param>
    ''' <returns></returns>
    Public Function loadBuyingInvocies1(ByVal orderIDParam As Integer) As DataSet
        Dim invoicesDataSet = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("SELECT  *
                                                    FROM    (				SELECT 
													InvoiceNumbers,
													ROW_NUMBER() OVER (PARTITION BY InvoiceNumbers, deliveryMergedID ORDER BY InvoiceNumbers, deliveryMergedID) AS RowNumber ,
													Names,
													deliveryID,
	                                                deliveryMergedID,
                                                    finalized,
                                                    exported,
													supplierCode
	                                                
                                                FROM   
                                                   (SELECT 
													
                                                    deliveryID,
	                                                deliveryMergedID,
                                                    finalized,
                                                    exported,
													supplierCode,
	                                                suppliersinvoices.supplierInvoiceNumber, 
													(CASE
													WHEN (suppliersinvoices.generalSupplier1InvoiceNumber = '????INV') THEN '????INV-1' ELSE suppliersinvoices.generalSupplier1InvoiceNumber END) as generalSupplier1InvoiceNumber,	 
													(CASE
													WHEN (suppliersinvoices.generalSupplier2InvoiceNumber = '????INV') THEN '????INV-2' ELSE suppliersinvoices.generalSupplier2InvoiceNumber END) as
	                                                generalSupplier2InvoiceNumber, 
													(CASE
													WHEN (suppliersinvoices.generalSupplier3InvoiceNumber = '????INV') THEN '????INV-3' ELSE suppliersinvoices.generalSupplier3InvoiceNumber END) as
	                                                generalSupplier3InvoiceNumber, 
													(CASE
													WHEN (suppliersinvoices.forwarderInvoiceNumber = '????INV') THEN '????INV-4' ELSE suppliersinvoices.forwarderInvoiceNumber END) as
	                                                forwarderInvoiceNumber, 
													(CASE
													WHEN (suppliersinvoices.clearingAgentInvoiceNumber = '????INV') THEN '????INV-5' ELSE suppliersinvoices.clearingAgentInvoiceNumber END) as
	                                                clearingAgentInvoiceNumber  
												
                                                FROM easyenquiry.suppliersinvoices 
                                                LEFT JOIN easyenquiry.delivery ON suppliersinvoices.delivery_ID = delivery.deliveryID
                                                LEFT JOIN easyenquiry.deliverymerged ON delivery.deliveryID = deliverymerged.delivery_ID
                                                LEFT JOIN easyenquiry.invoices ON deliverymerged.deliveryMergedID = invoices.deliveryMerged_ID
												LEFT JOIN easyenquiry.supplierquotations ON delivery.supQuotation_ID = supplierquotations.supQuotationID
												LEFT JOIN easyenquiry.suppliers ON supplierquotations.supplier_ID = suppliers.supplierID
                                                WHERE order_ID= '" & orderIDParam & "'
												) p 
                                                UNPIVOT  
                                                   (InvoiceNumbers 
	                                                FOR Names
	                                                IN   
                                                      (supplierInvoiceNumber,
	                                                  generalSupplier1InvoiceNumber, 
	                                                  generalSupplier2InvoiceNumber, 
	                                                  generalSupplier3InvoiceNumber, 
	                                                  forwarderInvoiceNumber, 
	                                                  clearingAgentInvoiceNumber)  
	                                                )AS 
	                                                unpvt 
                                                WHERE 
	                                                InvoiceNumbers <> '') AS a WHERE   a.RowNumber = 1
                                                ORDER BY deliveryID ASC, Names DESC", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(invoicesDataSet, "invoicesBuying")
        Return invoicesDataSet
    End Function

    ''' <summary>
    ''' LIST THE ITEMS OF EACH NUMBER FOR CALCULATIONS BUYING
    ''' </summary>
    Public Function loadItemsBuying(ByVal orderIDParam As Integer, ByVal mergedID As Integer, ByVal supInvNum As String) As DataSet
        ' Dataset and Datatable assign
        Dim invoicesDataSet = New DataSet

        Try
            ' Connection and Sql String
            Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
            Dim daFeatures As New SqlDataAdapter("SELECT distinct 
                                    quotItemListID,
                                    delivery.packingBuying,
                                    orders.orderID,
                                    quotations.quotationID,
                                   	deliverymerged.deliveryMergedID,
                                    deliveryitems.quotItemList_ID,
                                    deliveryReadyQty,
									deliveryitems.deliveryDeliveredQty,
                                    quotitemslist.quotItemPrice,
                                    suppliersinvoices.supplierInvoiceNumber,
                                    supplierquotations.supplier_ID,
                                    supplierquotations.supplierDiscountPer,
                                    supplierquotations.supplierDiscountLS,  
                                    supplierquotations.supplierCurrencyCode

                                    FROM easyenquiry.invoices

                                    LEFT JOIN easyenquiry.deliverymerged ON invoices.deliveryMerged_ID = deliverymerged.deliveryMergedID
									LEFT JOIN easyenquiry.delivery on deliverymerged.delivery_ID=deliveryID
                                    LEFT JOIN easyenquiry.orders ON delivery.order_ID = orders.orderID
									LEFT JOIN easyenquiry.quotations ON orders.quotation_ID = quotations.quotationID                                   
                                    LEFT JOIN easyenquiry.deliveryitems ON deliverymerged.delivery_ID = deliveryitems.delivery_ID
                                    LEFT JOIN easyenquiry.quotitemslist ON deliveryitems.quotItemList_ID = quotitemslist.quotItemListID
                                    LEFT JOIN easyenquiry.suppliersinvoices ON delivery.deliveryID = suppliersinvoices.delivery_ID
									LEFT JOIN easyenquiry.supplierquotations ON delivery.supQuotation_ID = supplierquotations.supQuotationID

                                    WHERE orderID = '" & orderIDParam & "'  AND deliveryMergedID = '" & mergedID & "' AND suppliersinvoices.supplierInvoiceNumber = '" & supInvNum & "'
                                    and quotItemPrice<>0
									ORDER BY quotItemListID ASC", cnDelivery)

            daFeatures.SelectCommand.CommandTimeout = 0

            ' data adapter filling dataset
            daFeatures.Fill(invoicesDataSet, "itemsBuying")

            daFeatures.Dispose()
        Catch ex As Exception
            MessageBox.Show("Load Items", "Importand Message", MessageBoxButtons.OK)
        End Try

        ' Returning datatable 
        Return invoicesDataSet
    End Function

    ''' <summary>
    ''' Get distinct values of supplier ID and invoice number
    ''' </summary>
    ''' <param name="orderIDParam"></param>
    ''' <returns></returns>
    Public Function loadSellingInvoices(ByVal orderIDParam As Integer) As DataSet
        Dim invoicesDataSet = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("SELECT DISTINCT
                                    
                                    invoices.invoiceNumber,
                                    invoices.deliveryMerged_ID,
                                    invoices.delivery_ID,
                                    invoices.invoiceType_ID,
                                    delivery.customerVat_ID,
                                    exported,
                                    finalized,
                                    (SELECT exchangerates.exchangeRateValue FROM EASYENQUIRY.EASYENQUIRY.exchangerates WHERE exchangerates.exchangeRateCurCode = customerCurrencyCode AND exchangerates.exchangeRateDate = quotations.quotationDate) as customerExchangeRate

                                    FROM easyenquiry.invoices

								    LEFT JOIN easyenquiry.delivery ON invoices.delivery_ID = delivery.deliveryID
                                    LEFT JOIN easyenquiry.orders ON delivery.order_ID = orders.orderID
                                    LEFT JOIN easyenquiry.quotations ON orders.quotation_ID = quotations.quotationID
                                    
                                    LEFT JOIN easyenquiry.deliverymerged ON invoices.deliveryMerged_ID = deliverymerged.deliveryMergedID
                                    LEFT JOIN easyenquiry.deliveryitems ON deliverymerged.deliveryMergedID = deliveryitems.deliveryMerged_ID
                                    LEFT JOIN easyenquiry.quotitemslist ON deliveryitems.quotItemList_ID = quotitemslist.quotItemListID
                                    LEFT JOIN easyenquiry.suppliersinvoices ON delivery.deliveryID = suppliersinvoices.delivery_ID
                                    LEFT JOIN easyenquiry.supplierquotations ON delivery.supQuotation_ID = supplierquotations.supQuotationID

                                    WHERE orderID = '" & orderIDParam & "'
                                    ORDER BY deliveryMerged_ID ASC", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(invoicesDataSet, "invoicesSelling")

        Return invoicesDataSet
    End Function

    ''' <summary>
    ''' Items Selling DataSet
    ''' </summary>
    Public Function loadItemsSelling(ByVal orderIDParam As Integer, Optional ByVal mergedID As String = "", Optional ByVal invNum As String = "") As DataSet
        ' Dataset and Datatable assign
        Dim invoicesDataSet = New DataSet
        Dim SQL = ""

        Try
            ' Connection and Sql String
            Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
            If mergedID = "" And invNum = "" Then
                SQL = "SELECT distinct
                                    invoices.invoiceID,
                                    orders.orderID,
                                    quotations.quotationID,                                   
                                    enquiries.enquiryID,
                                    deliverymerged.deliveryMergedID,
                                    quotations.customerCurrencyCode,
                                    deliveryitems.deliveryItemID,
                                    deliveryitems.quotItemList_ID,
                                    deliveryitems.deliveryDeliveredQty,
                                    quotitemslist.quotItemPrice,
                                    quotitemslist.quotItemIncrement,
                                    invoices.invoiceNumber,
                                    delivery.customerDiscountPer,
                                    delivery.customerDiscountLS,
																		(CASE
													WHEN (suppliersinvoices.supplierInvoiceNumber Is Null) THEN '0' ELSE suppliersinvoices.supplierInvoiceNumber END) as supplierInvoiceNumber


                                    FROM  easyenquiry.easyenquiry.invoices
                                    LEFT JOIN  easyenquiry.easyenquiry.deliverymerged ON invoices.deliveryMerged_ID = deliverymerged.deliveryMergedID
                                    LEFT JOIN  easyenquiry.easyenquiry.delivery ON deliverymerged.delivery_ID = delivery.deliveryID
                                    LEFT JOIN  easyenquiry.easyenquiry.orders ON delivery.order_ID = orders.orderID
                                    LEFT JOIN  easyenquiry.easyenquiry.quotations ON orders.quotation_ID = quotations.quotationID
                                    LEFT JOIN  easyenquiry.easyenquiry.enquiries ON quotations.enquiry_ID = enquiries.enquiryID
                                    LEFT JOIN  easyenquiry.easyenquiry.deliveryitems ON deliverymerged.deliveryMergedID = deliveryitems.deliveryMerged_ID
                                    LEFT JOIN  easyenquiry.easyenquiry.quotitemslist ON deliveryitems.quotItemList_ID = quotitemslist.quotItemListID
									LEFT JOIN easyenquiry.suppliersinvoices ON deliveryitems.delivery_ID = suppliersinvoices.delivery_ID

                                    WHERE orderID = '" & orderIDParam & "'
											AND deliveryDeliveredQty<>0
											AND quotItemPrice <>0"
            Else
                SQL = "SELECT distinct
                                    invoices.invoiceID,
                                    orders.orderID,
                                    quotations.quotationID,                                   
                                    enquiries.enquiryID,
                                    deliverymerged.deliveryMergedID,
                                    quotations.customerCurrencyCode,
                                    deliveryitems.deliveryItemID,
                                    deliveryitems.quotItemList_ID,
                                    deliveryitems.deliveryDeliveredQty,
                                    quotitemslist.quotItemPrice,
                                    quotitemslist.quotItemIncrement,
                                    invoices.invoiceNumber,
                                    delivery.customerDiscountPer,
                                    delivery.customerDiscountLS,
																		(CASE
													WHEN (suppliersinvoices.supplierInvoiceNumber Is Null) THEN '0' ELSE suppliersinvoices.supplierInvoiceNumber END) as supplierInvoiceNumber


                                    FROM  easyenquiry.easyenquiry.invoices
                                    LEFT JOIN  easyenquiry.easyenquiry.deliverymerged ON invoices.deliveryMerged_ID = deliverymerged.deliveryMergedID
                                    LEFT JOIN  easyenquiry.easyenquiry.delivery ON deliverymerged.delivery_ID = delivery.deliveryID
                                    LEFT JOIN  easyenquiry.easyenquiry.orders ON delivery.order_ID = orders.orderID
                                    LEFT JOIN  easyenquiry.easyenquiry.quotations ON orders.quotation_ID = quotations.quotationID
                                    LEFT JOIN  easyenquiry.easyenquiry.enquiries ON quotations.enquiry_ID = enquiries.enquiryID
                                    LEFT JOIN  easyenquiry.easyenquiry.deliveryitems ON deliverymerged.deliveryMergedID = deliveryitems.deliveryMerged_ID
                                    LEFT JOIN  easyenquiry.easyenquiry.quotitemslist ON deliveryitems.quotItemList_ID = quotitemslist.quotItemListID
									LEFT JOIN easyenquiry.suppliersinvoices ON deliveryitems.delivery_ID = suppliersinvoices.delivery_ID

                                    WHERE orderID = '" & orderIDParam & "'
                                            AND deliveryMergedID = '" & mergedID & "'
                                            AND invoiceNumber = '" & invNum & "'
											AND deliveryDeliveredQty<>0
											AND quotItemPrice <>0"
            End If
            Dim daFeatures As New SqlDataAdapter(SQL, cnDelivery)

            daFeatures.SelectCommand.CommandTimeout = 0

            ' data adapter filling dataset
            daFeatures.Fill(invoicesDataSet, "itemsSelling")

            daFeatures.Dispose()
        Catch ex As Exception
            MessageBox.Show("Load Items", "Important Message", MessageBoxButtons.OK)
        End Try
        ' Returning datatable 
        Return invoicesDataSet
    End Function

    Public Function loadExchangeRates(ByVal orderIDParam As Integer) As DataSet
        ' Dataset and Datatable assign
        Dim invoicesDataSet = New DataSet

        Try
            ' Connection and Sql String
            Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
            Dim daFeatures As New SqlDataAdapter("Select distinct
									supplierInvoiceNumber,
									(SELECT exchangerates.exchangeRateValue FROM EASYENQUIRY.EASYENQUIRY.exchangerates WHERE exchangerates.exchangeRateCurCode = supplierquotations.supplierCurrencyCode AND exchangerates.exchangeRateDate = quotations.quotationDate) as supplierExchangeRate,
									supplierquotations.supplierCurrencyCode
									
							from easyenquiry.suppliersinvoices
							LEFT JOIN easyenquiry.delivery ON suppliersinvoices.delivery_ID = delivery.deliveryID
							LEFT JOIN easyenquiry.supplierquotations ON delivery.supQuotation_ID = supplierquotations.supQuotationID
							LEFT JOIN easyenquiry.orders ON delivery.order_ID = orders.orderID
							LEFT JOIN easyenquiry.quotations ON orders.quotation_ID = quotations.quotationID
							where orderID = '" & orderIDParam & "'", cnDelivery)

            daFeatures.SelectCommand.CommandTimeout = 0

            ' data adapter filling dataset
            daFeatures.Fill(invoicesDataSet, "itemsSelling")

            daFeatures.Dispose()
        Catch ex As Exception
            MessageBox.Show("Load ExchangeRates", "Important Message", MessageBoxButtons.OK)
        End Try
        ' Returning datatable 
        Return invoicesDataSet
    End Function


    ''' <summary>
    ''' PACKING SELLING DATASET
    ''' </summary>
    Public Function loadPackingSelling(ByVal orderIDParam As Integer, ByVal mergedID As String, ByVal invNum As String) As DataSet
        ' Dataset and Datatable assign
        Dim invoicesDataSet = New DataSet

        Try
            ' Connection and Sql String
            Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
            Dim daFeatures As New SqlDataAdapter("SELECT distinct
                                    invoices.invoiceID,
                                    deliverymerged.deliveryMergedID,
                                    packingSelling,
                                    invoices.invoiceNumber

                                    FROM  easyenquiry.easyenquiry.invoices
                                    LEFT JOIN  easyenquiry.easyenquiry.deliverymerged ON invoices.deliveryMerged_ID = deliverymerged.deliveryMergedID
                                    LEFT JOIN  easyenquiry.easyenquiry.delivery ON deliverymerged.delivery_ID = delivery.deliveryID

                                    WHERE order_ID = '" & orderIDParam & "'
                                            AND deliveryMergedID = '" & mergedID & "'
                                            AND invoiceNumber = '" & invNum & "'", cnDelivery)

            daFeatures.SelectCommand.CommandTimeout = 0

            ' data adapter filling dataset
            daFeatures.Fill(invoicesDataSet, "itemsSelling")

            daFeatures.Dispose()
        Catch ex As Exception
            MessageBox.Show("Load Items", "Important Message", MessageBoxButtons.OK)
        End Try
        ' Returning datatable 
        Return invoicesDataSet
    End Function

    ''' <summary>
    ''' Freight, Extra Charges Buying Unpivot ******************* NOT USED YET ************
    ''' </summary>
    Public Function loadCharges(ByVal orderIDParam As Integer, ByVal mergedID As String, ByVal supInv As String) As DataSet
        ' Dataset and Datatable assign
        Dim invoicesDataSet = New DataSet

        Try
            ' Connection and Sql String
            Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
            Dim daFeatures As New SqlDataAdapter("SELECT 
										    * 
										    From (
										
										    SELECT
                                             InvoiceNumbers,
											 ROW_NUMBER() OVER (PARTITION BY InvoiceNumbers, deliveryMergedID ORDER BY InvoiceNumbers, deliveryMergedID) AS RowNumber,
									         Names,
                                            deliveryID,
                                             deliveryMergedID,
                                             supplierVat_ID,
                                             freightBuyingVat_ID,
                                             packingBuying,
                                             freightBuying,
                                             freightBuyingCurr,
                                             customsBuying,
                                             customsBuyingCurr,
                                             customsBuyingVat_ID,
                                             extraCharges1Buying,
                                             extraCharges1BuyingCurr,
                                             extraCharges1BuyingVat_ID,
                                             extraCharges2Buying,
                                             extraCharges2BuyingCurr,
                                             extraCharges2BuyingVat_ID,
                                             extraCharges3Buying,
                                             extraCharges3BuyingCurr,
                                             extraCharges3BuyingVat_ID

											From (

												SELECT
											supplierInvoiceNumber,
									        generalSupplier1InvoiceNumber,
									        generalSupplier2InvoiceNumber,
									        generalSupplier3InvoiceNumber,
									        forwarderInvoiceNumber,
									        clearingAgentInvoiceNumber,
                                            deliveryID,
                                            deliveryMergedID,
                                             supplierVat_ID,
                                             freightBuyingVat_ID,
                                             packingBuying,
                                             freightBuying,
                                             freightBuyingCurr,
                                             customsBuying,
                                             customsBuyingCurr,
                                             customsBuyingVat_ID,
                                             extraCharges1Buying,
                                             extraCharges1BuyingCurr,
                                             extraCharges1BuyingVat_ID,
                                             extraCharges2Buying,
                                             extraCharges2BuyingCurr,
                                             extraCharges2BuyingVat_ID,
                                             extraCharges3Buying,
                                             extraCharges3BuyingCurr,
                                             extraCharges3BuyingVat_ID

                                         FROM easyenquiry.delivery

                                            LEFT JOIN easyenquiry.suppliersinvoices ON delivery.deliveryID = suppliersinvoices.delivery_ID
                                            LEFT JOIN easyenquiry.orders ON delivery.order_ID = orders.orderID
                                            LEFT JOIN easyenquiry.deliverymerged ON delivery.deliveryID = deliverymerged.delivery_ID

                                         WHERE orderID = '" & orderIDParam & "'
                                            AND deliveryID = '" & mergedID & "'
                                            And supplierInvoiceNumber='" & supInv & "') p
										UNPIVOT
										(InvoiceNumbers FOR Names IN
										(
											supplierInvoiceNumber,
									        generalSupplier1InvoiceNumber,
									        generalSupplier2InvoiceNumber,
									        generalSupplier3InvoiceNumber,
									        forwarderInvoiceNumber,
									        clearingAgentInvoiceNumber
										
										
										))
										AS 
	                                                unpvt where InvoiceNumbers<>''
																
										) AS a WHERE   a.RowNumber = 1 
                                               ORDER BY deliveryID ASC, Names DESC", cnDelivery)

            daFeatures.SelectCommand.CommandTimeout = 0

            ' data adapter filling dataset
            daFeatures.Fill(invoicesDataSet, "extraCharges")

            daFeatures.Dispose()
        Catch ex As Exception
            MessageBox.Show("Load Charges", "Importand Message", MessageBoxButtons.OK)
        End Try
        ' Returning datatable 
        Return invoicesDataSet
    End Function

    ''' <summary>
    ''' Freight, Extra Charges Buying USED ONE
    ''' </summary>
    ''' <param name="orderIDParam"></param>
    ''' <param name="mergedID"></param>
    ''' <returns></returns>
    Public Function loadCharges1(ByVal orderIDParam As Integer, ByVal mergedID As String) As DataSet
        ' Dataset and Datatable assign
        Dim invoicesDataSet = New DataSet

        Try
            ' Connection and Sql String
            Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
            Dim daFeatures As New SqlDataAdapter("SELECT DISTINCT
                                            suppliersinvoices.supplierInvoiceNumber,
                                            generalSupplier1InvoiceNumber,
									        generalSupplier2InvoiceNumber,
									        generalSupplier3InvoiceNumber,
									        forwarderInvoiceNumber,
									        clearingAgentInvoiceNumber,
                                            deliverymerged.deliveryMergedID,
                                            deliveryID,
                                            delivery.supplierVat_ID,
                                            delivery.freightBuyingVat_ID,
                                            delivery.packingBuying,
                                            delivery.freightBuying,
                                            delivery.freightBuyingCurr,
                                            delivery.customsBuying,
                                            delivery.customsBuyingCurr,
                                            delivery.customsBuyingVat_ID,
                                            delivery.extraCharges1Buying,
                                            delivery.extraCharges1BuyingCurr,
                                            delivery.extraCharges1BuyingVat_ID,
                                            delivery.extraCharges2Buying,
                                            delivery.extraCharges2BuyingCurr,
                                            delivery.extraCharges2BuyingVat_ID,
                                            delivery.extraCharges3Buying,
                                            delivery.extraCharges3BuyingCurr,
                                            delivery.extraCharges3BuyingVat_ID

                                            FROM easyenquiry.delivery

                                            LEFT JOIN easyenquiry.suppliersinvoices ON delivery.deliveryID = suppliersinvoices.delivery_ID
                                            LEFT JOIN easyenquiry.orders ON delivery.order_ID = orders.orderID
                                            LEFT JOIN easyenquiry.deliverymerged ON delivery.deliveryID = deliverymerged.delivery_ID

                                            WHERE orderID = '" & orderIDParam & "'
                                            AND deliveryMergedID = '" & mergedID & "' AND (generalSupplier1InvoiceNumber <> '' OR generalSupplier2InvoiceNumber <> '' OR generalSupplier3InvoiceNumber <> '' OR forwarderInvoiceNumber <> '' OR clearingAgentInvoiceNumber <> '')", cnDelivery)

            daFeatures.SelectCommand.CommandTimeout = 0

            ' data adapter filling dataset
            daFeatures.Fill(invoicesDataSet, "extraCharges")

            daFeatures.Dispose()
        Catch ex As Exception
            MessageBox.Show("Load Charges", "Importand Message", MessageBoxButtons.OK)
        End Try
        ' Returning datatable 
        Return invoicesDataSet
    End Function

    ''' <summary>
    ''' Freight, Extra Charges Selling
    ''' </summary>
    Public Function loadChargesSelling(ByVal orderIDParam As Integer, ByVal mergedID As Integer, ByVal invNum As String) As DataSet
        ' Dataset and Datatable assign
        Dim invoicesDataSet = New DataSet

        Try
            ' Connection and Sql String
            Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
            Dim daFeatures As New SqlDataAdapter("SELECT DISTINCT
											invoices.deliveryMerged_ID,
                                            invoices.invoiceNumber,
                                            delivery.customerVat_ID,
                                            delivery.freightSellingVat_ID,
                                            delivery.freightSelling,
                                            delivery.freightSellingCurr,
                                            delivery.customsSelling,
                                            delivery.customsSellingCurr,
                                            delivery.customsSellingVat_ID,
                                            delivery.extraCharges1Selling,
                                            delivery.extraCharges1SellingCurr,
                                            delivery.extraCharges1SellingVat_ID,
                                            delivery.extraCharges2Selling,
                                            delivery.extraCharges2SellingCurr,
                                            delivery.extraCharges2SellingVat_ID,
                                            delivery.extraCharges3Selling,
                                            delivery.extraCharges3SellingCurr,
                                            delivery.extraCharges3SellingVat_ID

                                            FROM easyenquiry.easyenquiry.delivery

											LEFT JOIN easyenquiry.easyenquiry.orders ON delivery.order_ID = orders.orderID
											LEFT JOIN easyenquiry.easyenquiry.deliverymerged on deliverymerged.delivery_ID=deliveryID
											LEFT JOIN easyenquiry.easyenquiry.invoices On deliverymerged.deliveryMergedID = invoices.deliveryMerged_ID

                                            WHERE orderID = '" & orderIDParam & "'
                                            AND deliveryMerged_ID = '" & mergedID & "'
                                            AND invoiceNumber = '" & invNum & "'", cnDelivery)

            daFeatures.SelectCommand.CommandTimeout = 0

            ' data adapter filling dataset
            daFeatures.Fill(invoicesDataSet, "extraCharges")

            daFeatures.Dispose()
        Catch ex As Exception
            MessageBox.Show("Load Charges", "Importand Message", MessageBoxButtons.OK)
        End Try
        ' Returning datatable 
        Return invoicesDataSet
    End Function

    ''' <summary>
    ''' VAT Amount
    ''' </summary>
    Public Function loadVATPerc(ByVal vatIDParam As Integer) As Double
        ' Dataset and Datatable assign
        Dim invoicesDataSet = New DataSet

        Try
            ' Connection and Sql String
            Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
            Dim daFeatures As New SqlDataAdapter("SELECT 
                                            vat.vatID,
                                            vat.vatValue,
                                            vat.vatPos

                                            FROM easyenquiry.vat
                                            WHERE vatID = '" & vatIDParam & "'", cnDelivery)

            daFeatures.SelectCommand.CommandTimeout = 0

            ' data adapter filling dataset
            daFeatures.Fill(invoicesDataSet, "vat")

            daFeatures.Dispose()
        Catch ex As Exception
            MessageBox.Show("Load VAT", "Importand Message", MessageBoxButtons.OK)
        End Try

        Dim vatPerc As Double = 0.0
        If invoicesDataSet.Tables("vat").Rows.Count > 0 Then
            vatPerc = invoicesDataSet.Tables("vat").Rows(0).Item("vatValue")
        End If

        Return vatPerc
    End Function

    ''' <summary>
    ''' Get the rest of the data for the output
    ''' </summary>
    ''' <param name="orderIDParam"></param>
    ''' <param name="mergedIDParam"></param>
    ''' <returns></returns>
    Public Function loadOutputValues(ByVal orderIDParam As String, ByVal mergedIDParam As String, Optional ByVal deliveryIDParam As String = "") As DataSet
        ' Dataset and Datatable assign
        Dim ordersDataSet = New DataSet

        If deliveryIDParam = "" Then
            SQL = "SELECT  deliveryID, forwarderCompName, (SELECT clearingAgentCompName from [easyenquiry].[easyenquiry].clearingagents where clearingAgentID = clearingAgent_ID) AS clearingAgentCompName,
                                                (SELECT generalSupplierCompName from [easyenquiry].[easyenquiry].generalsuppliers where generalSupplierID= generalSupplier1_ID) AS generalSupplier1CompName,
	                                            (SELECT generalSupplierCompName from [easyenquiry].[easyenquiry].generalsuppliers where generalSupplierID= generalSupplier2_ID) AS generalSupplier2CompName,
		                                        (SELECT generalSupplierCompName from [easyenquiry].[easyenquiry].generalsuppliers where generalSupplierID= generalSupplier3_ID) AS generalSupplier3CompName,
                                                (SELECT forwarderCode FROM easyenquiry.easyenquiry.forwarders WHERE forwarderID = forwarder_ID) as forwarderCode,
                                                (SELECT generalSupplierCode FROM easyenquiry.easyenquiry.generalsuppliers WHERE generalSupplier1_ID = generalSupplierID) as generalSupplierID,
                                                COALESCE(enquiries.originalRefNum, enquiries.refNum) as originalRefNum,
                                                enquiries.enquiryID,
												processStatus_ID,
                                                paymentTermID, supplierCurrencyCode, hideVessel,enquiries.refNum, orders.orderID, invoices.invoiceType_ID, delivery.deliveryPlace, invoices.invoiceID, invoices.invoiceDate, suppliers.supplierCode, suppliers.supplierID ,supplierquotations.supplierPaymentTerm_ID, customers.customerCode, orders.custPurchOrderRef, delivery.dispatchDate, vessels.vesselName, '023' AS companyCode, ordercategories.purchaseCode, ordercategories.salesCode, invoices.invoiceNumber, suppliersinvoices.supplierInvoiceNumber, suppliersinvoices.supplierInvoiceDate, (SELECT incExpDanaosCodePur from easyenquiry.easyenquiry.incexpcatdanaos where vat.vatPos=navIncExpCatCode) as supIncExpDanaosCodePur, (SELECT incExpDanaosCodeSales from easyenquiry.easyenquiry.incexpcatdanaos LEFT JOIN easyenquiry.easyenquiry.vat on vat.vatID=enquiries.vat_ID where vat.vatPos=navIncExpCatCode) as cusIncExpDanaosCodePur, (SELECT incExpDanaosCodeSales from easyenquiry.easyenquiry.incexpcatdanaos LEFT JOIN easyenquiry.easyenquiry.vat on vat.vatID=enquiries.vat_ID where vat.vatPos=navIncExpCatCode) as cusIncExpDanaosCodeSales, (Select customerName from easyenquiry.easyenquiry.customers where enquiries.customer_ID=customerID)customerName, (Select supplierName from easyenquiry.easyenquiry.suppliers where supplierquotations.supplier_ID=supplierID)supplierName, (enquiries.refNum ) as narativeofentry, ( DATEADD(day,paymentTermValueInDays, CONVERT(date,dispatchDate))) AS dueDate2, CONVERT(VARCHAR(10),CAST(case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'dispatchDate' then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'invoiceDate' then  invoiceDate else deliveryDate end  AS DATE),103) AS dispatchDate, 
												case when ((SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) > 0) then case when convert(date, CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) , case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'dispatchDate' then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'invoiceDate' then  invoiceDate else deliveryDate end )) AS DATE),103),103)< convert(date, CONVERT(VARCHAR(10),CAST(invoices.invoiceDate AS DATE),103),103) then CONVERT(VARCHAR(10),CAST(invoiceDate AS DATE),103) else CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) , case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'dispatchDate' then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'invoiceDate' then  invoiceDate else deliveryDate end)) AS DATE),103) end else CONVERT(VARCHAR(10),CAST(invoiceDate AS DATE),103) end AS dueDate, case when ((SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) > 0) then case when convert(date, CONVERT(VARCHAR(10), CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID), case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = paymentTermCountFrom then dispatchDate else suppliersinvoices.supplierInvoiceDate end )) AS DATE),103),103) < convert(date, CONVERT(VARCHAR(10),CAST(suppliersinvoices.supplierInvoiceDate AS DATE),103),103) then CONVERT(VARCHAR(10),CAST(suppliersinvoices.supplierInvoiceDate AS DATE),103) else CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) , case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = paymentTermCountFrom then dispatchDate else suppliersinvoices.supplierInvoiceDate end )) AS DATE),103) end else CONVERT(VARCHAR(10),CAST(suppliersinvoices.supplierInvoiceDate AS DATE),103) end AS supplierDueDate,
												suppliersinvoices.supplierInvoiceDate, deliveryDate,paymentTermValueInDays, customerCurrencyCode FROM easyenquiry.easyenquiry.invoices LEFT JOIN easyenquiry.easyenquiry.deliverymerged ON invoices.deliveryMerged_ID=deliverymerged.deliveryMergedID LEFT JOIN easyenquiry.easyenquiry.delivery on deliverymerged.delivery_ID= delivery.deliveryID LEFT JOIN easyenquiry.easyenquiry.suppliersinvoices on suppliersinvoices.delivery_ID=delivery.deliveryID LEFT JOIN easyenquiry.easyenquiry.orders on delivery.order_ID=orders.orderID LEFT JOIN easyenquiry.easyenquiry.quotations ON orders.quotation_ID= quotations.quotationID LEFT JOIN easyenquiry.easyenquiry.enquiries ON quotations.enquiry_ID = enquiries.enquiryID LEFT JOIN easyenquiry.easyenquiry.supplierquotations ON supplierquotations.supQuotationID= delivery.supQuotation_ID LEFT JOIN easyenquiry.easyenquiry.suppliers ON suppliers.supplierID =supplierquotations.supplier_ID LEFT JOIN easyenquiry.easyenquiry.ordercategories ON enquiries.orderCategory_Code=ordercategories.orderCategoryCode LEFT JOIN easyenquiry.easyenquiry.paymentterms ON paymentterms.paymentTermID=quotations.paymentTerm_ID LEFT JOIN easyenquiry.easyenquiry.customers ON enquiries.customer_ID=customers.customerID LEFT JOIN easyenquiry.easyenquiry.vessels ON enquiries.vessel_ID=vessels.vesselID LEFT JOIN easyenquiry.easyenquiry.vat ON vat.vatID=supplierquotations.vat_ID LEFT JOIN [easyenquiry].easyenquiry.forwarders ON delivery.forwarder_ID = forwarders.forwarderID WHERE orderID='" & orderIDParam & "' AND deliveryMerged_ID = '" & mergedIDParam & "'"
        Else
            'SQL = "SELECT  deliveryID, forwarderCompName, (SELECT clearingAgentCompName from [easyenquiry].[easyenquiry].clearingagents where clearingAgentID = clearingAgent_ID) AS clearingAgentCompName,(SELECT generalSupplierCompName from [easyenquiry].[easyenquiry].generalsuppliers where generalSupplierID= generalSupplier1_ID) AS generalSupplier1CompName,(SELECT generalSupplierCompName from [easyenquiry].[easyenquiry].generalsuppliers where generalSupplierID= generalSupplier2_ID) AS generalSupplier2CompName,(SELECT generalSupplierCompName from [easyenquiry].[easyenquiry].generalsuppliers where generalSupplierID= generalSupplier3_ID) AS generalSupplier3CompName,(SELECT forwarderCode FROM easyenquiry.easyenquiry.forwarders WHERE forwarderID = forwarder_ID) as forwarderCode,(SELECT generalSupplierCode FROM easyenquiry.easyenquiry.generalsuppliers WHERE generalSupplier1_ID = generalSupplierID) as generalSupplierID,COALESCE(enquiries.originalRefNum, enquiries.refNum) as originalRefNum,enquiries.enquiryID,processStatus_ID,paymentTermID, supplierCurrencyCode, hideVessel,enquiries.refNum, orders.orderID, invoices.invoiceType_ID, delivery.deliveryPlace, invoices.invoiceID, invoices.invoiceDate, suppliers.supplierCode, suppliers.supplierID ,supplierquotations.supplierPaymentTerm_ID, customers.customerCode, orders.custPurchOrderRef, delivery.dispatchDate, vessels.vesselName, '023' AS companyCode, ordercategories.purchaseCode, ordercategories.salesCode, invoices.invoiceNumber, suppliersinvoices.supplierInvoiceNumber, suppliersinvoices.supplierInvoiceDate, (SELECT incExpDanaosCodePur from easyenquiry.easyenquiry.incexpcatdanaos where vat.vatPos=navIncExpCatCode) as supIncExpDanaosCodePur, (SELECT incExpDanaosCodeSales from easyenquiry.easyenquiry.incexpcatdanaos LEFT JOIN easyenquiry.easyenquiry.vat on vat.vatID=enquiries.vat_ID where vat.vatPos=navIncExpCatCode) as cusIncExpDanaosCodePur, (SELECT incExpDanaosCodeSales from easyenquiry.easyenquiry.incexpcatdanaos LEFT JOIN easyenquiry.easyenquiry.vat on vat.vatID=enquiries.vat_ID where vat.vatPos=navIncExpCatCode) as cusIncExpDanaosCodeSales, (Select customerName from easyenquiry.easyenquiry.customers where enquiries.customer_ID=customerID)customerName, (Select supplierName from easyenquiry.easyenquiry.suppliers where supplierquotations.supplier_ID=supplierID)supplierName, (enquiries.refNum ) as narativeofentry, ( DATEADD(day,paymentTermValueInDays, CONVERT(date,dispatchDate))) AS dueDate2, CONVERT(VARCHAR(10),CAST(case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'dispatchDate' then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'invoiceDate' then  invoiceDate else deliveryDate end  AS DATE),103) AS dispatchDate, case when ((SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) > 0) then case when convert(date, CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) , case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'dispatchDate' then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'invoiceDate' then  invoiceDate else deliveryDate end )) AS DATE),103),103)< convert(date, CONVERT(VARCHAR(10),CAST(invoices.invoiceDate AS DATE),103),103) then CONVERT(VARCHAR(10),CAST(invoiceDate AS DATE),103) else CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) , case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'dispatchDate' then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'invoiceDate' then  invoiceDate else deliveryDate end)) AS DATE),103) end else CONVERT(VARCHAR(10),CAST(invoiceDate AS DATE),103) end AS dueDate, case when ((SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) > 0) then case when convert(date, CONVERT(VARCHAR(10), CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID), case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = paymentTermCountFrom then dispatchDate else suppliersinvoices.supplierInvoiceDate end )) AS DATE),103),103) < convert(date, CONVERT(VARCHAR(10),CAST(suppliersinvoices.supplierInvoiceDate AS DATE),103),103) then CONVERT(VARCHAR(10),CAST(suppliersinvoices.supplierInvoiceDate AS DATE),103) else CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) , case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = paymentTermCountFrom then dispatchDate else suppliersinvoices.supplierInvoiceDate end )) AS DATE),103) end else CONVERT(VARCHAR(10),CAST(suppliersinvoices.supplierInvoiceDate AS DATE),103) end AS supplierDueDate,CONVERT(VARCHAR(10),CAST(case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =supplierquotations.supplierPaymentTerm_ID) = 'dispatchDate' then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =supplierquotations.supplierPaymentTerm_ID) = 'invoiceDate' then  invoiceDate else deliveryDate end  AS DATE),103) AS dispatchDate, case when ((SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =supplierquotations.supplierPaymentTerm_ID) > 0) then case when convert(date, CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =supplierquotations.supplierPaymentTerm_ID) , case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =supplierquotations.supplierPaymentTerm_ID) = 'dispatchDate' then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =supplierquotations.supplierPaymentTerm_ID) = 'invoiceDate' then  (select suppliersinvoices.supplierInvoiceDate from easyenquiry.easyenquiry.suppliersinvoices where delivery.deliveryID = suppliersinvoices.delivery_ID) else deliveryDate end )) AS DATE),103),103)< convert(date, CONVERT(VARCHAR(10),CAST((select suppliersinvoices.supplierInvoiceDate from easyenquiry.easyenquiry.suppliersinvoices where delivery.deliveryID = suppliersinvoices.delivery_ID) AS DATE),103),103) then CONVERT(VARCHAR(10),CAST(invoiceDate AS DATE),103) else CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =supplierquotations.supplierPaymentTerm_ID) , case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =supplierquotations.supplierPaymentTerm_ID) = 'dispatchDate' then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =supplierquotations.supplierPaymentTerm_ID) = 'invoiceDate' then  (select suppliersinvoices.supplierInvoiceDate from easyenquiry.easyenquiry.suppliersinvoices where delivery.deliveryID = suppliersinvoices.delivery_ID) else deliveryDate end)) AS DATE),103) end else CONVERT(VARCHAR(10),CAST((select suppliersinvoices.supplierInvoiceDate from easyenquiry.easyenquiry.suppliersinvoices where delivery.deliveryID = suppliersinvoices.delivery_ID) AS DATE),103) end AS supplierDueDate2 ,suppliersinvoices.supplierInvoiceDate, deliveryDate,paymentTermValueInDays, customerCurrencyCode FROM easyenquiry.easyenquiry.invoices LEFT JOIN easyenquiry.easyenquiry.deliverymerged ON invoices.deliveryMerged_ID=deliverymerged.deliveryMergedID LEFT JOIN easyenquiry.easyenquiry.delivery on deliverymerged.delivery_ID= delivery.deliveryID LEFT JOIN easyenquiry.easyenquiry.suppliersinvoices on suppliersinvoices.delivery_ID=delivery.deliveryID LEFT JOIN easyenquiry.easyenquiry.orders on delivery.order_ID=orders.orderID LEFT JOIN easyenquiry.easyenquiry.quotations ON orders.quotation_ID= quotations.quotationID LEFT JOIN easyenquiry.easyenquiry.enquiries ON quotations.enquiry_ID = enquiries.enquiryID LEFT JOIN easyenquiry.easyenquiry.supplierquotations ON supplierquotations.supQuotationID= delivery.supQuotation_ID LEFT JOIN easyenquiry.easyenquiry.suppliers ON suppliers.supplierID =supplierquotations.supplier_ID LEFT JOIN easyenquiry.easyenquiry.ordercategories ON enquiries.orderCategory_Code=ordercategories.orderCategoryCode LEFT JOIN easyenquiry.easyenquiry.paymentterms ON paymentterms.paymentTermID=quotations.paymentTerm_ID LEFT JOIN easyenquiry.easyenquiry.customers ON enquiries.customer_ID=customers.customerID LEFT JOIN easyenquiry.easyenquiry.vessels ON enquiries.vessel_ID=vessels.vesselID LEFT JOIN easyenquiry.easyenquiry.vat ON vat.vatID=supplierquotations.vat_ID LEFT JOIN [easyenquiry].easyenquiry.forwarders ON delivery.forwarder_ID = forwarders.forwarderID WHERE orderID='" & orderIDParam & "' AND deliveryMerged_ID = '" & mergedIDParam & "' and deliveryID = '" & deliveryIDParam & "'"
            SQL = "SELECT  deliveryID, forwarderCompName, (SELECT clearingAgentCompName from [easyenquiry].[easyenquiry].clearingagents where clearingAgentID = clearingAgent_ID) AS clearingAgentCompName,
                                                (SELECT generalSupplierCompName from [easyenquiry].[easyenquiry].generalsuppliers where generalSupplierID= generalSupplier1_ID) AS generalSupplier1CompName,
                                             (SELECT generalSupplierCompName from [easyenquiry].[easyenquiry].generalsuppliers where generalSupplierID= generalSupplier2_ID) AS generalSupplier2CompName,
                                          (SELECT generalSupplierCompName from [easyenquiry].[easyenquiry].generalsuppliers where generalSupplierID= generalSupplier3_ID) AS generalSupplier3CompName,
                                                (SELECT forwarderCode FROM easyenquiry.easyenquiry.forwarders WHERE forwarderID = forwarder_ID) as forwarderCode,
                                                (SELECT generalSupplierCode FROM easyenquiry.easyenquiry.generalsuppliers WHERE generalSupplier1_ID = generalSupplierID) as generalSupplierID,
                                                COALESCE(enquiries.originalRefNum, enquiries.refNum) as originalRefNum,
                                                enquiries.enquiryID,
            processStatus_ID,
                                                paymentTermID, supplierCurrencyCode, hideVessel,enquiries.refNum, orders.orderID, invoices.invoiceType_ID, delivery.deliveryPlace, invoices.invoiceID, invoices.invoiceDate, suppliers.supplierCode, suppliers.supplierID ,supplierquotations.supplierPaymentTerm_ID, customers.customerCode, orders.custPurchOrderRef, delivery.dispatchDate, vessels.vesselName, '023' AS companyCode, ordercategories.purchaseCode, ordercategories.salesCode, invoices.invoiceNumber, suppliersinvoices.supplierInvoiceNumber, suppliersinvoices.supplierInvoiceDate, (SELECT incExpDanaosCodePur from easyenquiry.easyenquiry.incexpcatdanaos where vat.vatPos=navIncExpCatCode) as supIncExpDanaosCodePur, (SELECT incExpDanaosCodeSales from easyenquiry.easyenquiry.incexpcatdanaos LEFT JOIN easyenquiry.easyenquiry.vat on vat.vatID=enquiries.vat_ID where vat.vatPos=navIncExpCatCode) as cusIncExpDanaosCodePur, (SELECT incExpDanaosCodeSales from easyenquiry.easyenquiry.incexpcatdanaos LEFT JOIN easyenquiry.easyenquiry.vat on vat.vatID=enquiries.vat_ID where vat.vatPos=navIncExpCatCode) as cusIncExpDanaosCodeSales, (Select customerName from easyenquiry.easyenquiry.customers where enquiries.customer_ID=customerID)customerName, (Select supplierName from easyenquiry.easyenquiry.suppliers where supplierquotations.supplier_ID=supplierID)supplierName, (enquiries.refNum ) as narativeofentry, ( DATEADD(day,paymentTermValueInDays, CONVERT(date,dispatchDate))) AS dueDate2, CONVERT(VARCHAR(10),CAST(case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'dispatchDate' then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'invoiceDate' then  invoiceDate else deliveryDate end  AS DATE),103) AS dispatchDate, 
            case when ((SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) > 0) then case when convert(date, CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) , case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'dispatchDate' then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'invoiceDate' then  invoiceDate else deliveryDate end )) AS DATE),103),103)< convert(date, CONVERT(VARCHAR(10),CAST(invoices.invoiceDate AS DATE),103),103) then CONVERT(VARCHAR(10),CAST(invoiceDate AS DATE),103) else CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) , case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'dispatchDate' then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'invoiceDate' then  invoiceDate else deliveryDate end)) AS DATE),103) end else CONVERT(VARCHAR(10),CAST(invoiceDate AS DATE),103) end AS dueDate, case when ((SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) > 0) then case when convert(date, CONVERT(VARCHAR(10), CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID), case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = paymentTermCountFrom then dispatchDate else suppliersinvoices.supplierInvoiceDate end )) AS DATE),103),103) < convert(date, CONVERT(VARCHAR(10),CAST(suppliersinvoices.supplierInvoiceDate AS DATE),103),103) then CONVERT(VARCHAR(10),CAST(suppliersinvoices.supplierInvoiceDate AS DATE),103) else CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) , case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = paymentTermCountFrom then dispatchDate else suppliersinvoices.supplierInvoiceDate end )) AS DATE),103) end else CONVERT(VARCHAR(10),CAST(suppliersinvoices.supplierInvoiceDate AS DATE),103) end AS supplierDueDate,
            CONVERT(VARCHAR(10),CAST(case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =supplierquotations.supplierPaymentTerm_ID) = 'dispatchDate' 
            then dispatchDate 
            when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =supplierquotations.supplierPaymentTerm_ID) = 'invoiceDate' 
            then  invoiceDate 
            else deliveryDate end  AS DATE),103) AS dispatchDate, 
            case when ((SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = supplierquotations.supplierPaymentTerm_ID) > 0) 
            then case when convert(date, CONVERT(VARCHAR(10), CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = supplierquotations.supplierPaymentTerm_ID), 
            case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = supplierquotations.supplierPaymentTerm_ID) = paymentTermCountFrom 
            then dispatchDate else suppliersinvoices.supplierInvoiceDate end )) AS DATE),103),103) 
            < 
            convert(date, CONVERT(VARCHAR(10),CAST(suppliersinvoices.supplierInvoiceDate AS DATE),103),103) 
            then CONVERT(VARCHAR(10),CAST(suppliersinvoices.supplierInvoiceDate AS DATE),103) else CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = supplierquotations.supplierPaymentTerm_ID) , case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = supplierquotations.supplierPaymentTerm_ID) = paymentTermCountFrom then dispatchDate else suppliersinvoices.supplierInvoiceDate end )) AS DATE),103) end else CONVERT(VARCHAR(10),CAST(suppliersinvoices.supplierInvoiceDate AS DATE),103) end AS supplierDueDate2,
            suppliersinvoices.supplierInvoiceDate, deliveryDate,paymentTermValueInDays, customerCurrencyCode FROM easyenquiry.easyenquiry.invoices LEFT JOIN easyenquiry.easyenquiry.deliverymerged ON invoices.deliveryMerged_ID=deliverymerged.deliveryMergedID LEFT JOIN easyenquiry.easyenquiry.delivery on deliverymerged.delivery_ID= delivery.deliveryID LEFT JOIN easyenquiry.easyenquiry.suppliersinvoices on suppliersinvoices.delivery_ID=delivery.deliveryID LEFT JOIN easyenquiry.easyenquiry.orders on delivery.order_ID=orders.orderID LEFT JOIN easyenquiry.easyenquiry.quotations ON orders.quotation_ID= quotations.quotationID LEFT JOIN easyenquiry.easyenquiry.enquiries ON quotations.enquiry_ID = enquiries.enquiryID LEFT JOIN easyenquiry.easyenquiry.supplierquotations ON supplierquotations.supQuotationID= delivery.supQuotation_ID LEFT JOIN easyenquiry.easyenquiry.suppliers ON suppliers.supplierID =supplierquotations.supplier_ID LEFT JOIN easyenquiry.easyenquiry.ordercategories ON enquiries.orderCategory_Code=ordercategories.orderCategoryCode LEFT JOIN easyenquiry.easyenquiry.paymentterms ON paymentterms.paymentTermID=quotations.paymentTerm_ID LEFT JOIN easyenquiry.easyenquiry.customers ON enquiries.customer_ID=customers.customerID LEFT JOIN easyenquiry.easyenquiry.vessels ON enquiries.vessel_ID=vessels.vesselID LEFT JOIN easyenquiry.easyenquiry.vat ON vat.vatID=supplierquotations.vat_ID LEFT JOIN [easyenquiry].easyenquiry.forwarders ON delivery.forwarder_ID = forwarders.forwarderID WHERE orderID='" & orderIDParam & "' AND deliveryMerged_ID = '" & mergedIDParam & "' and deliveryID = '" & deliveryIDParam & "'"

        End If

        Try
            ' Connection and Sql String
            Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
            Dim daFeatures As New SqlDataAdapter(SQL, cnDelivery)

            daFeatures.SelectCommand.CommandTimeout = 0

            ' data adapter filling dataset
            daFeatures.Fill(ordersDataSet, "outputvalues")

            daFeatures.Dispose()
        Catch ex As Exception
            MessageBox.Show("Load Output", "Importand Message", MessageBoxButtons.OK)
        End Try


        ' Returning datatable 
        Return ordersDataSet
    End Function

    Public Function ifExistsMS(ByVal deliveryID As Integer) As DataSet
        ' Dataset and Datatable assign
        Dim ordersDataSet = New DataSet

        Try
            ' Connection and Sql String
            Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
            Dim daFeatures As New SqlDataAdapter("  SELECT distinct 
		                                        supplierCode,
		                                        customerCode,
		                                        forwarderCode,
		                                        (SELECT generalSupplierCode from [easyenquiry].generalsuppliers where generalSupplierID= generalSupplier1_ID) AS generalSupplier1Code,
	                                            (SELECT generalSupplierCode from [easyenquiry].generalsuppliers where generalSupplierID= generalSupplier2_ID) AS generalSupplier2Code,
		                                        (SELECT generalSupplierCode from [easyenquiry].generalsuppliers where generalSupplierID= generalSupplier3_ID) AS generalSupplier3Code
		
                                          FROM easyenquiry.easyenquiry.delivery
                                          LEFT JOIN easyenquiry.easyenquiry.forwarders ON forwarder_ID = forwarderID
                                          LEFT JOIN easyenquiry.easyenquiry.generalsuppliers ON generalSupplier1_ID = generalSupplierID OR generalSupplier2_ID = generalSupplierID OR generalSupplier3_ID = generalSupplierID
                                          LEFT JOIN easyenquiry.easyenquiry.supplierquotations ON supQuotation_ID = supQuotationID
                                          LEFT JOIN easyenquiry.easyenquiry.suppliers ON supplier_ID = supplierID
                                          LEFT JOIN easyenquiry.easyenquiry.orders ON order_ID = orderID
                                          LEFT JOIN easyenquiry.easyenquiry.quotations ON quotation_ID = quotationID
                                          LEFT JOIN easyenquiry.easyenquiry.enquiries ON quotations.enquiry_ID = enquiryID
                                          LEFT JOIN easyenquiry.easyenquiry.customers ON customer_ID = customerID

                                          WHERE deliveryID = '" & deliveryID & "'", cnDelivery)

            daFeatures.SelectCommand.CommandTimeout = 0

            ' data adapter filling dataset
            daFeatures.Fill(ordersDataSet, "existsMS")

            daFeatures.Dispose()
        Catch ex As Exception
            MessageBox.Show("Load Code", "Importand Message", MessageBoxButtons.OK)
        End Try

        ' Returning datatable 
        Return ordersDataSet
    End Function

    Public Function ifExistsNames(ByVal deliveryID As Integer) As DataSet
        ' Dataset and Datatable assign
        Dim ordersDataSet = New DataSet

        Try
            ' Connection and Sql String
            Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
            Dim daFeatures As New SqlDataAdapter("  SELECT distinct 'Supplier Name :'+ supplierName as supplierName ,
		                                    'Customer Name :'+ customerName as customerName ,
		                                    'Forwarder Name :'+ forwarderName as forwarderName ,
		                                    'General Supplier 1 Name :'+ (SELECT generalSupplierName from [easyenquiry].generalsuppliers where generalSupplierID= generalSupplier1_ID) as generalSupplierName ,
		                                    'General Supplier 2 Name :'+ (SELECT generalSupplierName from [easyenquiry].generalsuppliers where generalSupplierID= generalSupplier2_ID) as generalSupplierName ,
		                                    'General Supplier 3 Name :'+ (SELECT generalSupplierName from [easyenquiry].generalsuppliers where generalSupplierID= generalSupplier3_ID) as generalSupplierName 
		
                                      FROM easyenquiry.easyenquiry.delivery
                                      LEFT JOIN easyenquiry.easyenquiry.forwarders ON forwarder_ID = forwarderID
                                      LEFT JOIN easyenquiry.easyenquiry.generalsuppliers ON generalSupplier1_ID = generalSupplierID OR generalSupplier2_ID = generalSupplierID OR generalSupplier3_ID = generalSupplierID
                                      LEFT JOIN easyenquiry.easyenquiry.supplierquotations ON supQuotation_ID = supQuotationID
                                      LEFT JOIN easyenquiry.easyenquiry.suppliers ON supplier_ID = supplierID
                                      LEFT JOIN easyenquiry.easyenquiry.orders ON order_ID = orderID
                                      LEFT JOIN easyenquiry.easyenquiry.quotations ON quotation_ID = quotationID
                                      LEFT JOIN easyenquiry.easyenquiry.enquiries ON quotations.enquiry_ID = enquiryID
                                      LEFT JOIN easyenquiry.easyenquiry.customers ON customer_ID = customerID

                                      WHERE deliveryID = '" & deliveryID & "'", cnDelivery)

            daFeatures.SelectCommand.CommandTimeout = 0

            ' data adapter filling dataset
            daFeatures.Fill(ordersDataSet, "existsMS")

            daFeatures.Dispose()
        Catch ex As Exception
            MessageBox.Show("Load Code", "Importand Message", MessageBoxButtons.OK)
        End Try

        ' Returning datatable 
        Return ordersDataSet
    End Function

    ''' <summary>
    ''' CREDIT NOTE TYPE 
    ''' </summary>
    ''' <param name="orderID"></param>
    ''' <returns></returns>
    Public Function creditNoteType(ByVal orderID As Integer) As DataSet
        ' Dataset and Datatable assign
        Dim ordersDataSet = New DataSet

        Try
            ' Connection and Sql String
            Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
            Dim daFeatures As New SqlDataAdapter(" SELECT Distinct deliverymerged.deliveryMergedID,
                                                delivery.deliveryID,
		                                        creditNoteType_ID,
		                                        CreditNoteTypeName,
                                                creditnotes.exported,
                                                creditnotes.creditNoteNumber,
                                                creditnotes.creditNoteDate,
CONVERT(VARCHAR(10),CAST(case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'dispatchDate' then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =quotations.paymentTerm_ID) = 'invoiceDate' then creditNoteDate else deliveryDate end  AS DATE),103) AS dispatchDate, case when ((SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) > 0) then case when convert(date, CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) , case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =quotations.paymentTerm_ID) = 'dispatchDate' then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =quotations.paymentTerm_ID) = 'invoiceDate' then creditNoteDate else deliveryDate end )) AS DATE),103),103)< convert(date, CONVERT(VARCHAR(10),CAST(creditnotes.creditNoteDate AS DATE),103),103) then CONVERT(VARCHAR(10),CAST(creditNoteDate AS DATE),103) else CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) , case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'dispatchDate' then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'invoiceDate' then creditNoteDate  else deliveryDate end)) AS DATE),103) end else CONVERT(VARCHAR(10),CAST(creditNoteDate  AS DATE),103) end AS dueDate2
                                                

		                                        FROM easyenquiry.creditnotes 
		                                        LEFT JOIN easyenquiry.creditnotetypes ON creditnotes.creditNoteType_ID = creditNoteTypeID
		                                        LEFT JOIN easyenquiry.delivery ON creditnotes.delivery_ID = delivery.deliveryID
                                                LEFT JOIN easyenquiry.deliverymerged ON delivery.deliveryID = deliverymerged.delivery_ID
                                                LEFT JOIN easyenquiry.invoices ON invoices.deliveryMerged_ID = deliverymerged.deliveryMergedID
                                                LEFT JOIN easyenquiry.orders ON delivery.order_ID = orders.orderID
                                                LEFT JOIN easyenquiry.quotations ON orders.quotation_ID = quotations.quotationID

		                                        WHERE order_ID = '" & orderID & "' and creditnotes.exported = 0", cnDelivery)


            daFeatures.SelectCommand.CommandTimeout = 0

            ' data adapter filling dataset
            daFeatures.Fill(ordersDataSet, "creditnote")

            daFeatures.Dispose()
        Catch ex As Exception
            MessageBox.Show("Load Credit", "Importand Message", MessageBoxButtons.OK)
        End Try

        ' Returning datatable 
        Return ordersDataSet
    End Function

    ''' <summary>
    ''' CREDIT NOTE BROKER
    ''' </summary>
    ''' <param name="orderIDParam"></param>
    ''' <param name="deliveryIDParam"></param>
    ''' <returns></returns>
    Public Function creditNoteBroker(ByVal orderIDParam As Integer, ByVal deliveryMergedIDParam As Integer, Optional ByVal deliveryBrokerID As Integer = 0) As DataSet
        ' Dataset and Datatable assign
        Dim ordersDataSet = New DataSet

        Try
            ' Connection and Sql String
            Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
            Dim daFeatures As New SqlDataAdapter("
SELECT orderID,
                                            deliverymerged.deliveryMergedID,
		                                    delivery.deliveryID,
		                                    deliverybrokers.deliveryBrokerageFeeLS,
		                                    deliverybrokers.deliveryBrokerageFeePer,
		                                    brokerageCurrencyCode,
		                                    deliveryBrokerID,
                                            brokerCode,
									(SELECT exchangerates.exchangeRateValue FROM EASYENQUIRY.EASYENQUIRY.exchangerates WHERE exchangerates.exchangeRateCurCode = quotationsbrokers.brokerageCurrencyCode AND exchangerates.exchangeRateDate = quotations.quotationDate) as brokerageExchangeRate


                                            FROM easyenquiry.quotationsbrokers

LEFT JOIN EASYENQUIRY.EASYENQUIRY.quotations ON quotationsbrokers.quotation_ID=quotations.quotationID 
LEFT JOIN EASYENQUIRY.EASYENQUIRY.orders ON quotations.quotationID=orders.quotation_ID 
LEFT JOIN EASYENQUIRY.EASYENQUIRY.delivery ON orders.orderID=delivery.order_ID 
LEFT JOIN EASYENQUIRY.EASYENQUIRY.deliverybrokers ON deliverybrokers.quotationBroker_ID=quotationsbrokers.quotationBrokerID 
LEFT JOIN easyenquiry.deliverymerged ON deliverybrokers.delivery_ID = deliverymerged.delivery_ID
LEFT JOIN EASYENQUIRY.EASYENQUIRY.brokers ON quotationsbrokers.broker_ID=brokers.brokerID 
                                            WHERE order_ID = '" & orderIDParam & "' AND deliveryMergedID ='" & deliveryMergedIDParam & "' AND deliveryBrokerID = '" & deliveryBrokerID & "'", cnDelivery)


            daFeatures.SelectCommand.CommandTimeout = 0

            ' data adapter filling dataset
            daFeatures.Fill(ordersDataSet, "creditnote")

            daFeatures.Dispose()
        Catch ex As Exception
            MessageBox.Show("Load Credit", "Importand Message", MessageBoxButtons.OK)
        End Try

        ' Returning datatable 
        Return ordersDataSet
    End Function

    ''' <summary>
    ''' CREDIT NOTE CUSTOMER
    ''' </summary>
    ''' <param name="orderIDParam"></param>
    ''' <param name="deliveryMergedIDParam"></param>
    ''' <returns></returns>
    Public Function creditNoteCustomer(ByVal orderIDParam As Integer, ByVal deliveryMergedIDParam As Integer) As DataSet
        ' Dataset and Datatable assign
        Dim ordersDataSet = New DataSet

        Try
            ' Connection and Sql String
            Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
            Dim daFeatures As New SqlDataAdapter("SELECT order_ID,
                                            deliverymerged.deliveryMergedID,
		                                    delivery.deliveryID,
		                                    creditnotes.creditNoteCustLSAmount,
		                                    creditnotes.creditNoteDiscountLS,
		                                    creditNoteNumber,
		                                    creditNoteID,
		                                    creditNoteDelQty,
		                                    cnCustItemPrice

		                                    FROM easyenquiry.creditnotes
		                                    LEFT JOIN easyenquiry.delivery ON creditnotes.delivery_ID = delivery.deliveryID
		                                    LEFT JOIN easyenquiry.creditnotedeliveryitems ON creditnotes.creditNoteID = creditNote_ID
                                            LEFT JOIN easyenquiry.deliverymerged ON delivery.deliveryID = deliverymerged.delivery_ID

		                                    WHERE order_ID = '" & orderIDParam & "' AND deliveryMergedID ='" & deliveryMergedIDParam & "' and (creditnotes.creditNoteCustLSAmount <> 0 OR creditnotes.creditNoteDiscountLS <> 0) and creditNoteNumber <> ''", cnDelivery)


            daFeatures.SelectCommand.CommandTimeout = 0

            ' data adapter filling dataset
            daFeatures.Fill(ordersDataSet, "creditnote")

            daFeatures.Dispose()
        Catch ex As Exception
            MessageBox.Show("Load Credit", "Importand Message", MessageBoxButtons.OK)
        End Try

        ' Returning datatable 
        Return ordersDataSet
    End Function

    ''' <summary>
    '''  Credit Note supplier
    ''' </summary>
    ''' <param name="orderIDParam"></param>
    ''' <param name="deliveryMergedIDParam"></param>
    ''' <returns></returns>
    Public Function creditNoteSupplier(ByVal orderIDParam As Integer, ByVal deliveryMergedIDParam As Integer) As DataSet
        ' Dataset and Datatable assign
        Dim ordersDataSet = New DataSet

        Try
            ' Connection and Sql String
            Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
            Dim daFeatures As New SqlDataAdapter("SELECT DISTINCT order_ID,
                                            deliverymerged.deliveryMergedID,
		                                    delivery.deliveryID,
											creditNoteSupDate,
		                                    creditnotes.creditNoteSupNumber,
		                                    creditnotes.creditNoteSupLSAmount,
                                            cnForwFreightPrice,
                                            cnForwNumber,
                                            cnForwDate,
                                            cnGenSup1Number,
                                            cnGenSupExtChar1Price,
                                            freightBuyingCurr,
                                            extraCharges1BuyingCurr
                                            
		                                    FROM easyenquiry.creditnotes
		                                    LEFT JOIN easyenquiry.delivery ON creditnotes.delivery_ID = delivery.deliveryID
                                            LEFT JOIN easyenquiry.deliverymerged ON delivery.deliveryID = deliverymerged.delivery_ID

		                                    WHERE order_ID = '" & orderIDParam & "' AND deliveryMergedID ='" & deliveryMergedIDParam & "' AND (creditNoteSupLSAmount <> 0 OR CnForwFreightPrice <> 0 OR cnGenSupExtChar1Price <> 0)", cnDelivery)


            daFeatures.SelectCommand.CommandTimeout = 0

            ' data adapter filling dataset
            daFeatures.Fill(ordersDataSet, "creditnote")

            daFeatures.Dispose()
        Catch ex As Exception
            MessageBox.Show("Load Credit", "Importand Message", MessageBoxButtons.OK)
        End Try

        ' Returning datatable 
        Return ordersDataSet
    End Function

    ''' <summary>
    ''' DEBIT NOTE TYPE ******************* OLD NOT USED **************************
    ''' </summary>
    ''' <param name="orderID"></param>
    ''' <returns></returns>
    Public Function debitNoteType1(ByVal orderID As Integer) As DataSet
        ' Dataset and Datatable assign
        Dim ordersDataSet = New DataSet

        Try
            ' Connection and Sql String
            Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
            Dim daFeatures As New SqlDataAdapter("SELECT deliverymerged.deliveryMergedID,
                                                delivery.deliveryID,
												debitNoteID,
												debitNoteNumber,
		                                        debitNoteType_ID,
                                                debitNoteCreatedDate,
                                                debitNoteDate
                                                
		                                        FROM easyenquiry.debitnotes 

		                                        LEFT JOIN easyenquiry.delivery ON debitnotes.delivery_ID = delivery.deliveryID
                                                LEFT JOIN easyenquiry.deliverymerged ON delivery.deliveryID = deliverymerged.delivery_ID

		                                        WHERE order_ID = '" & orderID & "'", cnDelivery)


            daFeatures.SelectCommand.CommandTimeout = 0

            ' data adapter filling dataset
            daFeatures.Fill(ordersDataSet, "debitnote")

            daFeatures.Dispose()
        Catch ex As Exception
            MessageBox.Show("Load Debit", "Importand Message", MessageBoxButtons.OK)
        End Try

        ' Returning datatable 
        Return ordersDataSet
    End Function

    ''' <summary>
    ''' DEBIT NOTE TYPE - NEW USED - DUE DATE UPDATED
    ''' </summary>
    ''' <param name="orderID"></param>
    ''' <returns></returns>
    Public Function debitNoteType(ByVal orderID As Integer) As DataSet
        ' Dataset and Datatable assign
        Dim ordersDataSet = New DataSet

        Try
            ' Connection and Sql String
            Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
            Dim daFeatures As New SqlDataAdapter("SELECT deliverymerged.deliveryMergedID,
                                                delivery.deliveryID,
												debitNoteID,
												debitNoteNumber,
		                                        debitNoteType_ID,
                                                debitNoteCreatedDate,
                                                debitNoteDate,
exported,
quotations.paymentTerm_ID,
 CONVERT(VARCHAR(10),CAST(case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'dispatchDate' then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =quotations.paymentTerm_ID) = 'invoiceDate' then  debitNoteDate else deliveryDate end  AS DATE),103) AS dispatchDate, case when ((SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) > 0) then case when convert(date, CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) , case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =quotations.paymentTerm_ID) = 'dispatchDate' then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =quotations.paymentTerm_ID) = 'invoiceDate' then debitNoteDate else deliveryDate end )) AS DATE),103),103)< convert(date, CONVERT(VARCHAR(10),CAST(debitnotes.debitNoteDate AS DATE),103),103) then CONVERT(VARCHAR(10),CAST(debitNoteDate AS DATE),103) else CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) , case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'dispatchDate' then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'invoiceDate' then debitNoteDate else deliveryDate end)) AS DATE),103) end else CONVERT(VARCHAR(10),CAST(debitNoteDate AS DATE),103) end AS dueDate2

		                                        FROM easyenquiry.debitnotes 

		                                        LEFT JOIN easyenquiry.delivery ON debitnotes.delivery_ID = delivery.deliveryID
                                                LEFT JOIN easyenquiry.deliverymerged ON delivery.deliveryID = deliverymerged.delivery_ID
                                                LEFT JOIN easyenquiry.orders ON delivery.order_ID = orders.orderID
                                                LEFT JOIN easyenquiry.quotations ON orders.quotation_ID = quotations.quotationID

		                                        WHERE order_ID = '" & orderID & "'", cnDelivery)


            daFeatures.SelectCommand.CommandTimeout = 0

            ' data adapter filling dataset
            daFeatures.Fill(ordersDataSet, "debitnote")

            daFeatures.Dispose()
        Catch ex As Exception
            MessageBox.Show("Load Debit", "Importand Message", MessageBoxButtons.OK)
        End Try

        ' Returning datatable 
        Return ordersDataSet
    End Function

    ''' <summary>
    ''' DEBIT NOTE BROKER
    ''' </summary>
    ''' <param name="orderIDParam"></param>
    ''' <param name="deliveryIDParam"></param>
    ''' <returns></returns>
    Public Function debitNoteBroker(ByVal orderIDParam As Integer, ByVal deliveryMergedIDParam As Integer) As DataSet
        ' Dataset and Datatable assign
        Dim ordersDataSet = New DataSet

        Try
            ' Connection and Sql String
            Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
            Dim daFeatures As New SqlDataAdapter("SELECT order_ID,
                                              deliveryID,
                                              deliveryMergedID,
                                              debitNoteID,
                                              debitNoteNumber,
                                              debitNoteBrokerLS,
                                              brokerageCurrencyCode,
                                              brokerCode,
                                              debitNoteDate

                                              FROM easyenquiry.debitnotes

		                                      LEFT JOIN easyenquiry.deliverybrokers ON debitnotes.delivery_ID = deliverybrokers.delivery_ID
		                                      LEFT JOIN easyenquiry.quotationsbrokers ON deliverybrokers.quotationBroker_ID = quotationsbrokers.quotationBrokerID
		                                      LEFT JOIN easyenquiry.brokers ON quotationsbrokers.broker_ID = brokers.brokerID
                                              LEFT JOIN easyenquiry.delivery ON delivery.deliveryID = debitnotes.delivery_ID
                                              LEFT JOIN easyenquiry.deliverymerged ON deliverymerged.delivery_ID = delivery.deliveryID

                                              WHERE order_ID = '" & orderIDParam & "' AND deliveryMergedID = '" & deliveryMergedIDParam & "'", cnDelivery)


            daFeatures.SelectCommand.CommandTimeout = 0

            ' data adapter filling dataset
            daFeatures.Fill(ordersDataSet, "debitnote")

            daFeatures.Dispose()
        Catch ex As Exception
            MessageBox.Show("Load Debit", "Importand Message", MessageBoxButtons.OK)
        End Try

        ' Returning datatable 
        Return ordersDataSet
    End Function

    ''' <summary>
    ''' DEBIT NOTE CUSTOMER
    ''' </summary>
    ''' <param name="orderIDParam"></param>
    ''' <param name="deliveryMergedIDParam"></param>
    ''' <returns></returns>
    Public Function debitNoteCustomer(ByVal orderIDParam As Integer, ByVal deliveryMergedIDParam As Integer) As DataSet
        ' Dataset and Datatable assign
        Dim ordersDataSet = New DataSet

        Try
            ' Connection and Sql String
            Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
            Dim daFeatures As New SqlDataAdapter("SELECT order_ID,
                                              deliveryID,
                                              deliveryMergedID,
                                              debitNoteID,
                                              debitNoteNumber,
											  debitNoteItemSupplierType_ID,
                                              (SELECT forwarderCode FROM easyenquiry.easyenquiry.forwarders WHERE forwarderID = debitNoteItemSupplierID) as forwarderCode,
											  (SELECT clearingAgentCode FROM easyenquiry.easyenquiry.clearingagents WHERE clearingAgentID = debitNoteItemSupplierID) as clearingAgentCode,
                                              debitNoteSellingCurrencyCode,
                                              debitNoteItemBuyingValue,
                                              debitNoteItemBuyingVat_ID,
                                              debitNoteItemSellingValue,
                                              debitNoteItemSellingVat_ID,
                                              debitNoteItemQuantity,
                                                debitNoteDate,
                                              debitNoteItemSupplierInvoiceDate,
                                              debitNoteItemSupplierInvoiceNumber

                                              FROM easyenquiry.debitnotes

                                              LEFT JOIN easyenquiry.delivery ON delivery.deliveryID = debitnotes.delivery_ID
                                              LEFT JOIN easyenquiry.deliverymerged ON deliverymerged.delivery_ID = delivery.deliveryID
                                              LEFT JOIN easyenquiry.debitnotesitems ON debitnotesitems.debitNote_ID = debitnotes.debitNoteID

                                              WHERE order_ID = '" & orderIDParam & "' AND deliveryMergedID = '" & deliveryMergedIDParam & "'", cnDelivery)


            daFeatures.SelectCommand.CommandTimeout = 0

            ' data adapter filling dataset
            daFeatures.Fill(ordersDataSet, "debitnote")

            daFeatures.Dispose()
        Catch ex As Exception
            MessageBox.Show("Load Debit", "Importand Message", MessageBoxButtons.OK)
        End Try

        ' Returning datatable 
        Return ordersDataSet
    End Function

    ''' <summary>
    ''' Create Credit Note Broker in frmDeliveryFind for the dropdown box toi display all brokers
    ''' </summary>
    ''' <param name="orderIDParam"></param>
    ''' <returns></returns>
    Public Function createCreditNoteBroker(ByVal orderIDParam As Integer) As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("begin
                                        Declare @vlakas INT = (Select Count(deliveryBrokerID) From easyenquiry.easyenquiry.deliverybrokers left join easyenquiry.easyenquiry.DELIVERY on delivery_ID = deliveryID where deliveryID = deliverybrokers.delivery_ID and order_ID = '" & orderIDParam & "')
                                        DECLARE @piovlakas INT = (Select Count(creditNoteID) From easyenquiry.easyenquiry.creditnotes left join easyenquiry.easyenquiry.delivery on delivery_ID = deliveryID where deliveryID = creditnotes.delivery_ID and order_ID = '" & orderIDParam & "')

                                        --set vlakas  (Select Count(creditNoteID) From easyenquiry.creditnotes where deliveryID = creditnotes.delivery_ID)

                                        select                                  

                                        @piovlakas as vlakas,
                                        @vlakas as piovlakas,
                                        deliveryBrokerID,
                                        delivery.deliveryID,
                                        quotationBroker_ID,
                                        deliveryBrokerageFeePer,
                                        deliveryBrokerageFeeLS,
                                        creditNoteID,
                                        brokerName,
                                        brokerCompName
                                        
                                    From easyenquiry.easyenquiry.quotationsbrokers

                                    Left Join easyenquiry.easyenquiry.deliverybrokers On quotationsbrokers.quotationBrokerID = deliverybrokers.quotationBroker_ID
                                    Left join easyenquiry.easyenquiry.delivery on deliverybrokers.delivery_ID = delivery.deliveryID
                                    left join easyenquiry.easyenquiry.creditnotes on deliveryID = creditnotes.delivery_ID
                                    Left join easyenquiry.easyenquiry.brokers On broker_ID = brokerID
                                    where order_ID = '" & orderIDParam & "' and (@vlakas<>@piovlakas or @piovlakas<>0) AND (DELIVERYBROKERS.deliveryBrokerID<>CREDITNOTES.deliveryBroker_ID Or CREDITNOTES.deliveryBroker_ID is Null)

                                    end", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(finalOrderSummary, "finalordersum")
        Return finalOrderSummary
    End Function

    ''' <summary>
    ''' Delivery Broker ID
    ''' </summary>
    ''' <param name="creditNoteID"></param>
    ''' <returns></returns>
    Public Function getDeliveryBrokerIDFromCNotes(ByVal creditNoteID As Integer) As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("select deliverybroker_ID 
                                        From easyenquiry.creditnotes
                                        where creditnoteID = '" & creditNoteID & "'", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(finalOrderSummary, "finalordersum")
        Return finalOrderSummary
    End Function

    Public Function cnCreatedDelBrokers(ByVal orderIDParam As Integer) As DataSet
        ' Dataset and Datatable assign
        Dim ordersDataSet = New DataSet

        Try
            ' Connection and Sql String
            Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
            Dim daFeatures As New SqlDataAdapter("SELECT deliveryBroker_ID

                                              FROM easyenquiry.easyenquiry.creditnotes

                                              LEFT JOIN easyenquiry.easyenquiry.delivery ON delivery.deliveryID = creditnotes.delivery_ID

                                              WHERE order_ID = '" & orderIDParam & "'", cnDelivery)


            daFeatures.SelectCommand.CommandTimeout = 0

            ' data adapter filling dataset
            daFeatures.Fill(ordersDataSet, "debitnote")

            daFeatures.Dispose()
        Catch ex As Exception
            MessageBox.Show("Load Debit", "Importand Message", MessageBoxButtons.OK)
        End Try

        ' Returning datatable 
        Return ordersDataSet
    End Function

    ''' <summary>
    ''' Get dates for invoices and credit notes to check if they match
    ''' </summary>
    ''' <param name="orderIDParam"></param>
    ''' <returns></returns>
    Public Function getDates(ByVal orderIDParam As Integer, Optional ByVal tmpCreditNote As String = "") As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        If tmpCreditNote <> "" Then
            Dim daFeatures As New SqlDataAdapter("SELECT Distinct deliveryMergedID,
				invoiceDate,			
				creditNoteDate,
                creditNoteType_ID
				FROM easyenquiry.delivery
				LEFT JOIN easyenquiry.creditnotes ON delivery.deliveryID = creditnotes.delivery_ID
				LEFT JOIN easyenquiry.deliverymerged ON delivery.deliveryID = deliverymerged.delivery_ID
				LEFT JOIN easyenquiry.invoices ON deliveryMergedID = invoices.deliveryMerged_ID
				where order_ID = '" & orderIDParam & "' and creditNoteID = '" & tmpCreditNote & "' and creditNoteDate Is Not Null
			", cnDelivery)

            daFeatures.SelectCommand.CommandTimeout = 0

            ' data adapter filling dataset
            daFeatures.Fill(finalOrderSummary, "finalordersum")
        Else
            Dim daFeatures As New SqlDataAdapter("SELECT Distinct deliveryMergedID,
				invoiceDate,			
				creditNoteDate,
                creditNoteType_ID
				FROM easyenquiry.delivery
				LEFT JOIN easyenquiry.creditnotes ON delivery.deliveryID = creditnotes.delivery_ID
				LEFT JOIN easyenquiry.deliverymerged ON delivery.deliveryID = deliverymerged.delivery_ID
				LEFT JOIN easyenquiry.invoices ON deliveryMergedID = invoices.deliveryMerged_ID
				where order_ID = '" & orderIDParam & "' And creditNoteDate Is Not Null
			", cnDelivery)
            daFeatures.SelectCommand.CommandTimeout = 0

            ' data adapter filling dataset
            daFeatures.Fill(finalOrderSummary, "finalordersum")
        End If

        Return finalOrderSummary
    End Function

    ''' <summary>
    ''' get supplier invoices to check if they are inserted correctly before export
    ''' </summary>
    ''' <param name="orderIDParam"></param>
    ''' <returns></returns>
    Public Function getSupplierInvoices(ByVal orderIDParam As Integer) As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("Select distinct supplierInvoiceNumber,
                                           deliveryID,
deliverymergedID
	                                       From easyenquiry.delivery
	                                       Left Join easyenquiry.suppliersinvoices on delivery.deliveryID = suppliersinvoices.delivery_ID
	                                       Left Join easyenquiry.deliverymerged on delivery.deliveryID = deliverymerged.delivery_ID
	                                       where order_ID =  '" & orderIDParam & "'", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(finalOrderSummary, "finalordersum")
        Return finalOrderSummary
    End Function

    ''' <summary>
    '''  CHeck if invoices are exported before the edit supplier invoices
    ''' </summary>
    ''' <param name="deliveryMergedIDParam"></param>
    ''' <returns></returns>
    Public Function checkInvoices(ByVal deliveryMergedIDParam As Integer) As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("Select invoiceID, exported from [easyenquiry].[easyenquiry].invoices
	                                       where deliveryMerged_ID =  '" & deliveryMergedIDParam & "'", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(finalOrderSummary, "finalordersum")
        Return finalOrderSummary
    End Function

    ''' <summary>
    ''' Get the data of each supplier invoice Delivery
    ''' </summary>
    ''' <param name="deliveryIDParam"></param>
    ''' <returns></returns>
    Public Function getSupInvoicesData(ByVal deliveryIDParam As Integer) As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("
SELECT [deliveryID]
      ,[order_ID]
	  ,(Select forwarderCompName From easyenquiry.forwarders where forwarder_ID = forwarderID) as forwarderCompName
      ,[freightBuying]
      ,[freightBuyingCurr]
	  ,(Select vatValue From easyenquiry.vat Where freightBuyingVat_ID = vatID) as freightVatPercentage
	  ,(Select clearingAgentCompName From easyenquiry.clearingagents where clearingAgent_ID = clearingAgentID) as clearingAgentCompName
      ,[customsBuying]
      ,[customsBuyingCurr]
      ,(Select vatValue From easyenquiry.vat Where customsBuyingVat_ID = vatID) as customsVatPercentage
	  ,(Select generalSupplierCompName From easyenquiry.generalsuppliers where generalSupplier1_ID = generalSupplierID) as generalSupplier1CompName
      ,[extraCharges1Buying]
      ,[extraCharges1BuyingCurr]
      ,(Select vatValue From easyenquiry.vat Where extraCharges1BuyingVat_ID = vatID) as genSup1VatPercentage
	  ,(Select generalSupplierCompName From easyenquiry.generalsuppliers where generalSupplier2_ID = generalSupplierID) as generalSupplier2CompName
      ,[extraCharges2Buying]
      ,[extraCharges2BuyingCurr]
      ,(Select vatValue From easyenquiry.vat Where extraCharges2BuyingVat_ID = vatID) as genSup2VatPercentage
	  ,(Select generalSupplierCompName From easyenquiry.generalsuppliers where generalSupplier3_ID = generalSupplierID) as generalSupplier3CompName
      ,[extraCharges3Buying]
      ,[extraCharges3BuyingCurr]
      ,(Select vatValue From easyenquiry.vat Where extraCharges3BuyingVat_ID = vatID) as genSup3VatPercentage
  FROM [easyenquiry].[easyenquiry].[delivery]
  where deliveryID = '" & deliveryIDParam & "'", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(finalOrderSummary, "finalordersum")
        Return finalOrderSummary
    End Function

    ''' <summary>
    ''' list the adjustment types in the combo box
    ''' </summary>
    ''' <param name="deliveryMergedIDParam"></param>
    ''' <returns></returns>
    Public Function adjustmentTypesQuery() As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("SELECT  [adjustmentTypeID]
      ,[adjustmentTypeName]
  FROM [easyenquiry].[easyenquiry].[adjustmenttypes]", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(finalOrderSummary, "finalordersum")
        Return finalOrderSummary
    End Function

    ''' <summary>
    ''' Show currencies
    ''' </summary>
    ''' <returns></returns>
    Public Function showCurrencies() As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("SELECT currencyCode
  FROM [easyenquiry].[easyenquiry].[currencies]", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(finalOrderSummary, "finalordersum")
        Return finalOrderSummary
    End Function

    ''' <summary>
    ''' ge tthe vat types
    ''' </summary>
    ''' <returns></returns>
    Public Function vatTypes() As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("SELECT  [vatID]
      ,[vatValue]
  FROM [easyenquiry].[easyenquiry].[vat]", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(finalOrderSummary, "finalordersum")
        Return finalOrderSummary
    End Function

    ''' <summary>
    '''  Check if adjustment exists
    ''' </summary>
    ''' <param name="adjID"></param>
    ''' <returns></returns>
    Public Function adjustmentsExist(ByVal adjID As String) As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("Select delivery_ID,
		                                        adjustmentID,
		                                        adjustmentDate,
		                                        (Select adjustmentTypeName from easyenquiry.adjustmenttypes where adjustments.adjustmentType_ID = adjustmentTypeID) as adjustmentTypeName,
		                                        forwarder_ID,
		                                        clearingAgent_ID,
		                                        generalSupplier1_ID,
		                                        generalSupplier2_ID,
		                                        generalSupplier3_ID,
                                                adjustmentVatLS
		                                        From easyenquiry.easyenquiry.adjustments where adjustmentID = '" & adjID & "' AND (forwarder_ID <> 0 Or clearingAgent_ID <> 0 Or generalSupplier1_ID <> 0 or generalSupplier2_ID <> 0 or generalSupplier3_ID <> 0)
                                        ", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(finalOrderSummary, "finalordersum")
        Return finalOrderSummary
    End Function

    ''' <summary>
    ''' Export adjustments
    ''' </summary>
    ''' <param name="orderID"></param>
    ''' <returns></returns>
    Public Function dsAdjustmentNote(ByVal orderID As String) As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("Select  deliveryMergedID,
		                                            deliveryID,
		                                            adjustmentID,
		                                            adjustmentNumber,
		                                            adjustmentDate,
		                                            adjustmentTypeName,
		                                            exported,
        											adjustments.forwarder_ID,
													adjustments.clearingAgent_ID,
													adjustments.generalSupplier1_ID,
													adjustments.generalSupplier2_ID,
													adjustments.generalSupplier3_ID,
                                                    adjustments.adjustmentVatLS,
														CONVERT(VARCHAR(10),CAST(case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'dispatchDate' then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =quotations.paymentTerm_ID) = 'invoiceDate' then  adjustmentDate else deliveryDate end  AS DATE),103) AS dispatchDate, 
													
													case when ((SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) > 0) 
													then 
													case when convert(date, CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) , 
													case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =quotations.paymentTerm_ID) = 'dispatchDate' 
													then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =quotations.paymentTerm_ID) = 'invoiceDate' 
													then suppliersinvoices.forwarderInvoiceDate else deliveryDate end )) AS DATE),103),103)
													< 
													convert(date, CONVERT(VARCHAR(10),CAST(suppliersinvoices.forwarderInvoiceDate AS DATE),103),103) 
													then CONVERT(VARCHAR(10),CAST(suppliersinvoices.forwarderInvoiceDate AS DATE),103) 
													else CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) , 
													case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'dispatchDate' 
													then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = quotations.paymentTerm_ID) = 'invoiceDate' 
													then suppliersinvoices.forwarderInvoiceDate else deliveryDate end)) AS DATE),103) end else CONVERT(VARCHAR(10),CAST(suppliersinvoices.forwarderInvoiceDate AS DATE),103) end AS dueDate2,
					
													case when ((SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.freightPaymentTerm_ID) > 0) 
													then 
													case when convert(date, CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.freightPaymentTerm_ID) , 
													case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =delivery.freightPaymentTerm_ID) = 'invoiceDate' 
													then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =delivery.freightPaymentTerm_ID) = 'dispatchDate' 
													then suppliersinvoices.forwarderInvoiceDate else deliveryDate end )) AS DATE),103),103)
													< 
													convert(date, CONVERT(VARCHAR(10),CAST(suppliersinvoices.forwarderInvoiceDate AS DATE),103),103) 
													then CONVERT(VARCHAR(10),CAST(suppliersinvoices.forwarderInvoiceDate AS DATE),103) 
													else CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.freightPaymentTerm_ID) , 
													case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.freightPaymentTerm_ID) = 'invoiceDate' 
													then dispatchDate 
													when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.freightPaymentTerm_ID) = 'dispatchDate' 
													then suppliersinvoices.forwarderInvoiceDate  else deliveryDate end)) AS DATE),103) end else CONVERT(VARCHAR(10),CAST(suppliersinvoices.forwarderInvoiceDate  AS DATE),103) end AS dueDateF,

													case when ((SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.generalSupplier1PaymentTerm_ID) > 0) 
													then 
													case when convert(date, CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.generalSupplier1PaymentTerm_ID) , 
													case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =delivery.generalSupplier1PaymentTerm_ID) = 'invoiceDate' 
													then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =delivery.generalSupplier1PaymentTerm_ID) = 'dispatchDate' 
													then suppliersinvoices.generalSupplier1InvoiceDate else deliveryDate end )) AS DATE),103),103)
													< 
													convert(date, CONVERT(VARCHAR(10),CAST(suppliersinvoices.generalSupplier1InvoiceDate AS DATE),103),103) 
													then CONVERT(VARCHAR(10),CAST(suppliersinvoices.generalSupplier1InvoiceDate AS DATE),103) 
													else CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.generalSupplier1PaymentTerm_ID) , 
													case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.generalSupplier1PaymentTerm_ID) = 'invoiceDate' 
													then dispatchDate 
													when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.generalSupplier1PaymentTerm_ID) = 'dispatchDate' 
													then suppliersinvoices.generalSupplier1InvoiceDate  else deliveryDate end)) AS DATE),103) end else CONVERT(VARCHAR(10),CAST(suppliersinvoices.generalSupplier1InvoiceDate  AS DATE),103) end AS dueDateGS1,

													case when ((SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.generalSupplier2PaymentTerm_ID) > 0) 
													then 
													case when convert(date, CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.generalSupplier2PaymentTerm_ID) , 
													case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =delivery.generalSupplier2PaymentTerm_ID) = 'invoiceDate' 
													then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =delivery.generalSupplier2PaymentTerm_ID) = 'dispatchDate' 
													then suppliersinvoices.generalSupplier2InvoiceDate else deliveryDate end )) AS DATE),103),103)
													< 
													convert(date, CONVERT(VARCHAR(10),CAST(suppliersinvoices.generalSupplier2InvoiceDate AS DATE),103),103) 
													then CONVERT(VARCHAR(10),CAST(suppliersinvoices.generalSupplier2InvoiceDate AS DATE),103) 
													else CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.generalSupplier2PaymentTerm_ID) , 
													case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.generalSupplier2PaymentTerm_ID) = 'invoiceDate' 
													then dispatchDate 
													when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.generalSupplier2PaymentTerm_ID) = 'dispatchDate' 
													then suppliersinvoices.generalSupplier2InvoiceDate  else deliveryDate end)) AS DATE),103) end else CONVERT(VARCHAR(10),CAST(suppliersinvoices.generalSupplier2InvoiceDate  AS DATE),103) end AS dueDateGS2,

													case when ((SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.generalSupplier3PaymentTerm_ID) > 0) 
													then 
													case when convert(date, CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.generalSupplier3PaymentTerm_ID) , 
													case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =delivery.generalSupplier3PaymentTerm_ID) = 'invoiceDate' 
													then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =delivery.generalSupplier3PaymentTerm_ID) = 'dispatchDate' 
													then suppliersinvoices.generalSupplier3InvoiceDate else deliveryDate end )) AS DATE),103),103)
													< 
													convert(date, CONVERT(VARCHAR(10),CAST(suppliersinvoices.generalSupplier3InvoiceDate AS DATE),103),103) 
													then CONVERT(VARCHAR(10),CAST(suppliersinvoices.generalSupplier3InvoiceDate AS DATE),103) 
													else CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.generalSupplier3PaymentTerm_ID) , 
													case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.generalSupplier3PaymentTerm_ID) = 'invoiceDate' 
													then dispatchDate 
													when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID = delivery.generalSupplier3PaymentTerm_ID) = 'dispatchDate' 
													then suppliersinvoices.generalSupplier3InvoiceDate  else deliveryDate end)) AS DATE),103) end else CONVERT(VARCHAR(10),CAST(suppliersinvoices.generalSupplier3InvoiceDate  AS DATE),103) end AS dueDateGS3

		                                            From easyenquiry.easyenquiry.adjustments 
		                                            LEFT JOIN easyenquiry.easyenquiry.adjustmenttypes ON adjustments.adjustmentType_ID = adjustmentTypeID
		                                            LEFT JOIN easyenquiry.easyenquiry.delivery ON adjustments.delivery_ID = deliveryID
		                                            LEFT JOIN easyenquiry.easyenquiry.suppliersinvoices ON delivery.deliveryID = suppliersinvoices.delivery_ID
	                                                LEFT JOIN easyenquiry.easyenquiry.deliverymerged ON deliveryID = deliverymerged.delivery_ID
													LEFT JOIN easyenquiry.easyenquiry.orders ON order_ID = orderID
													LEFT JOIN easyenquiry.easyenquiry.quotations ON quotation_ID = quotationID

		                                            where delivery.order_ID = '" & orderID & "'", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(finalOrderSummary, "finalordersum")
        Return finalOrderSummary
    End Function

    ''' <summary>
    ''' Get the invoices of each delivery
    ''' </summary>
    ''' <param name="deliveryID"></param>
    ''' <returns></returns>
    Public Function dsInvSuppliers(ByVal deliveryID As String) As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("SELECT
                                              [forwarderInvoiceNumber]
                                              ,[clearingAgentInvoiceNumber]
                                              ,[generalSupplier1InvoiceNumber]
                                              ,[generalSupplier2InvoiceNumber]
                                              ,[generalSupplier3InvoiceNumber]
                                          FROM [easyenquiry].[easyenquiry].[suppliersinvoices]
                                          where delivery_ID = '" & deliveryID & "'", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(finalOrderSummary, "finalordersum")
        Return finalOrderSummary
    End Function

    Public Function ifANExists(ByVal deliveryIDParam As String) As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("Select adjustmentID,
	                                            adjustmentNumber,
	                                            delivery_ID,
	                                            forwarder_ID,
	                                            clearingAgent_ID,
	                                            generalSupplier1_ID,
	                                            generalSupplier2_ID,
	                                            generalSupplier3_ID
	                                            From easyenquiry.easyenquiry.adjustments
	                                            where delivery_ID = '" & deliveryIDParam & "'", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(finalOrderSummary, "finalordersum")
        Return finalOrderSummary
    End Function

    ''' <summary>
    ''' ADjustment Note Summary data
    ''' </summary>
    ''' <param name="orderIDParam"></param>
    ''' <param name="deliveryIDParam"></param>
    ''' <param name="adjustmentIDParam"></param>
    ''' <returns></returns>
    Public Function reportAdjustmentDS(ByVal orderIDParam As String, ByVal deliveryIDParam As String, ByVal adjustmentIDParam As String) As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter(";WITH CTE AS ( 
SELECT DISTINCT delivery.deliveryID,
orders.custPurchOrderRef, 
orders.orderID, 
orders.vatValue, 
orderslist.orderListID, 
quotations.quotationID,
CONVERT(VARCHAR(10),CAST(orderDate AS DATE),104) AS orderDate, 
adminusers.userFullName, 
case when ((SELECT showPreSoftRefNumOnInvoice from easyenquiry.easyenquiry.config)=1 OR (orders.previousSoftRefNum !='')) then orders.previousSoftRefNum else enquiries.refNum end AS refNum, 
enquiries.customerRef, 
customers.customerName, 
customers.customerPostCode, 
customers.customerTown, 
customers.customerAddress1, 
customers.customerAddress2, 
customers.customerAddress3, 
(SELECT countries.countryName from easyenquiry.easyenquiry.countries WHERE countries.countryID=customers.country_ID) AS customerCountryName, 
(SELECT paymentterms.paymentTermName from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID=quotations.paymentTerm_ID) AS paymentTermName, 
vessels.vesselName, 
suppliers.supplierName, 
(SELECT paymentterms.paymentTermName from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID=supplierquotations.supplierPaymentTerm_ID) AS supplierPaymentTerm_ID,
ordercategories.orderCategoryName, 
invoices.invoiceNumber, 
CONVERT(VARCHAR(10),CAST(invoicedate AS DATE),104)AS invoiceDate, 
invoices.invoiceType_ID, 
invoices.finalized, 
enquiries.enquiryDesc,
delivery.supplierDiscountLS,
delivery.supplierDiscountPer,
supplierquotations.supplierCurrencyCode, 
orders.orderSummaryNotes, 
delivery.deliveryPlace, 
delivery.deliveryMethod_ID,
deliverymethods.deliveryMethodName, 
CONVERT(VARCHAR(10),CAST(dispatchDate AS DATE),104)AS dispatchDate, 
case when ((SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID=quotations.paymentTerm_ID) > 0) then CONVERT(VARCHAR(10),CAST( (SELECT dateadd(DAY, (SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID=quotations.paymentTerm_ID),dispatchDate ) ) as date),104) else 'PAID' end AS dueDate, 
agents.agentAddress, 
agents.agentTown, 
(SELECT countries.countryName from easyenquiry.easyenquiry.countries WHERE countries.countryID=agents.agentCountryID) AS agentCountryName,
delivery.freightBuying, 
delivery.freightBuyingCurr, 
(SELECT forwarderCompName FROM easyenquiry.easyenquiry.forwarders WHERE adjustments.forwarder_ID = forwarderID) as forwarderCompName,
(SELECT vatValue FROM easyenquiry.easyenquiry.vat WHERE freightBuyingVat_ID = vatID) as freightVat,
(SELECT forwarderAddress From easyenquiry.easyenquiry.forwarders where adjustments.forwarder_ID = forwarderID) as forwarderAddress,
(SELECT forwarderTown From easyenquiry.easyenquiry.forwarders where adjustments.forwarder_ID = forwarderID) as forwarderTown,
(SELECT countryName From easyenquiry.easyenquiry.forwarders LEFT jOIN easyenquiry.easyenquiry.countries on forwarderCountryID = countryID where adjustments.forwarder_ID = forwarderID) as forwarderCountryName,
delivery.customsBuying, 
delivery.customsBuyingCurr, 
(SELECT clearingAgentCompName FROM easyenquiry.easyenquiry.clearingagents WHERE adjustments.clearingAgent_ID = clearingAgentID) as clearingAgentCompName,
(SELECT vatValue FROM easyenquiry.easyenquiry.vat WHERE customsBuyingVat_ID = vatID) as customsVat,
(SELECT clearingAgentAddress From easyenquiry.easyenquiry.clearingagents where adjustments.clearingAgent_ID = clearingAgentID) as clearingAgentAddress,
(SELECT clearingAgentTown From easyenquiry.easyenquiry.clearingagents where adjustments.clearingAgent_ID = clearingAgentID) as clearingAgentTown,
(SELECT countryName From easyenquiry.easyenquiry.clearingagents LEFT jOIN easyenquiry.easyenquiry.countries on clearingAgentCountryID = countryID where adjustments.clearingAgent_ID = clearingAgentID) as clearingAgentCountryName,
delivery.extraCharges1Buying, 
delivery.extraCharges1BuyingCurr, 
(SELECT generalSupplierCompName FROM easyenquiry.easyenquiry.generalsuppliers WHERE adjustments.generalSupplier1_ID = generalSupplierID) as generalSupplier1CompName,
(SELECT vatValue FROM easyenquiry.easyenquiry.vat WHERE extraCharges1BuyingVat_ID = vatID) as extraCharges1Vat,
(SELECT generalSupplierAddress From easyenquiry.easyenquiry.generalsuppliers where adjustments.generalSupplier1_ID = generalSupplierID) as generalSupplier1Address,
(SELECT generalSupplierTown From easyenquiry.easyenquiry.generalsuppliers where adjustments.generalSupplier1_ID = generalSupplierID) as generalSupplier1Town,
(SELECT countryName From easyenquiry.easyenquiry.generalsuppliers LEFT jOIN easyenquiry.easyenquiry.countries on generalSupplierCountryID = countryID where adjustments.generalSupplier1_ID = generalSupplierID) as generalSupplier1CountryName,
delivery.extraCharges2Buying, 
delivery.extraCharges2BuyingCurr, 
(SELECT generalSupplierCompName FROM easyenquiry.easyenquiry.generalsuppliers WHERE adjustments.generalSupplier2_ID = generalSupplierID) as generalSupplier2CompName,
(SELECT vatValue FROM easyenquiry.easyenquiry.vat WHERE extraCharges2BuyingVat_ID = vatID) as extraCharges2Vat,
(SELECT generalSupplierAddress From easyenquiry.easyenquiry.generalsuppliers where adjustments.generalSupplier2_ID = generalSupplierID) as generalSupplier2Address,
(SELECT generalSupplierTown From easyenquiry.easyenquiry.generalsuppliers where adjustments.generalSupplier2_ID = generalSupplierID) as generalSupplier2Town,
(SELECT countryName From easyenquiry.easyenquiry.generalsuppliers LEFT jOIN easyenquiry.easyenquiry.countries on generalSupplierCountryID = countryID where adjustments.generalSupplier2_ID = generalSupplierID) as generalSupplier2CountryName,
delivery.extraCharges3Buying, 
delivery.extraCharges3BuyingCurr, 
(SELECT generalSupplierCompName FROM easyenquiry.easyenquiry.generalsuppliers WHERE adjustments.generalSupplier3_ID = generalSupplierID) as generalSupplier3CompName,
(SELECT vatValue FROM easyenquiry.easyenquiry.vat WHERE extraCharges3BuyingVat_ID = vatID) as extraCharges3Vat,
(SELECT generalSupplierAddress From easyenquiry.easyenquiry.generalsuppliers where adjustments.generalSupplier3_ID = generalSupplierID) as generalSupplier3Address,
(SELECT generalSupplierTown From easyenquiry.easyenquiry.generalsuppliers where adjustments.generalSupplier3_ID = generalSupplierID) as generalSupplier3Town,
(SELECT countryName From easyenquiry.easyenquiry.generalsuppliers LEFT jOIN easyenquiry.easyenquiry.countries on generalSupplierCountryID = countryID where adjustments.generalSupplier3_ID = generalSupplierID) as generalSupplier3CountryName,
(SELECT adjustmentTypeName FROM easyenquiry.easyenquiry.adjustmenttypes WHERE adjustmentType_ID = adjustmentTypeID) as adjustmentTypeName,
adjustments.adjustmentNumber,
adjustments.adjustmentDate


from easyenquiry.easyenquiry.delivery 
left join easyenquiry.easyenquiry.orders ON delivery.order_ID=orders.orderID 
left join easyenquiry.easyenquiry.orderslist ON orders.orderID=orderslist.order_ID 
left join easyenquiry.easyenquiry.quotations ON orders.quotation_ID=quotations.quotationID 
left join easyenquiry.easyenquiry.supplierevalquotation ON quotations.supEvalQuot_ID=supplierevalquotation.supEvalQuotID 
left join easyenquiry.easyenquiry.supplierevalquotationlist ON supplierevalquotation.supEvalQuotID=supplierevalquotationlist.supEvalQuot_ID 
left join easyenquiry.easyenquiry.supplierquotations ON supplierevalquotationlist.supQuotation_ID=supplierquotations.supQuotationID 
left join easyenquiry.easyenquiry.enquiries ON supplierevalquotation.enquiry_ID=enquiries.enquiryID 
left join easyenquiry.easyenquiry.adminusers ON enquiries.user_ID=adminusers.userID 
left join easyenquiry.easyenquiry.customers ON enquiries.customer_ID=customers.customerID 
left join easyenquiry.easyenquiry.suppliers ON supplierquotations.supplier_ID=suppliers.supplierID 
left join easyenquiry.easyenquiry.ordercategories ON enquiries.orderCategory_Code=ordercategories.orderCategoryCode 
left join easyenquiry.easyenquiry.vessels ON enquiries.vessel_ID=vessels.vesselID 
left join easyenquiry.easyenquiry.deliverymethods ON delivery.deliveryMethod_ID=deliverymethods.deliveryMethodID 
left join easyenquiry.easyenquiry.agents ON delivery.agent_ID=agents.agentID 
left join easyenquiry.easyenquiry.invoices ON delivery.deliveryID=invoices.delivery_ID 
LEFT JOIN easyenquiry.easyenquiry.adjustments ON delivery.deliveryID = adjustments.delivery_ID

WHERE orders.orderID='" & orderIDParam & "' 
AND delivery.deliveryID='" & deliveryIDParam & "' 
AND adjustments.adjustmentID = '" & adjustmentIDParam & "' ) 

SELECT deliveryID ,
deliveryPlace ,
custPurchOrderRef ,
orderID ,
vatValue ,
orderListID ,
quotationID ,
orderDate ,
userFullName ,
refNum ,
customerRef ,
customerName ,
customerPostCode ,
customerTown ,
customerAddress1 ,
customerAddress2 ,
customerAddress3 ,
customerCountryName ,
paymentTermName ,
vesselName ,
supplierName ,
supplierPaymentTerm_ID ,
orderCategoryName ,
invoiceNumber ,
invoiceDate ,
invoiceType_ID ,
finalized ,
enquiryDesc ,
supplierDiscountLS ,
supplierDiscountPer ,
supplierCurrencyCode ,
orderSummaryNotes ,
deliveryPlace ,
deliveryMethod_ID ,
deliveryMethodName ,
dispatchDate ,
dueDate ,
agentAddress ,
agentTown ,
agentCountryName,
freightBuying ,
freightBuyingCurr ,
forwarderCompName,
freightVat,
forwarderAddress,
forwarderTown,
forwarderCountryName,
customsBuying,
customsBuyingCurr,
clearingAgentCompName,
customsVat,
clearingAgentAddress,
clearingAgentTown,
clearingAgentCountryName,
extraCharges1Buying,
extraCharges1BuyingCurr,
generalSupplier1CompName,
extraCharges1Vat,
generalSupplier1Address,
generalSupplier1Town,
generalSupplier1CountryName,
extraCharges2Buying,
extraCharges2BuyingCurr,
generalSupplier2CompName,
extraCharges2Vat,
generalSupplier2Address,
generalSupplier2Town,
generalSupplier2CountryName,
extraCharges3Buying,
extraCharges3BuyingCurr,
generalSupplier3CompName,
extraCharges3Vat,
generalSupplier3Address,
generalSupplier3Town,
generalSupplier3CountryName,
adjustmentTypeName,
adjustmentNumber,
adjustmentDate

FROM cte

", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(finalOrderSummary, "finalordersum")
        Return finalOrderSummary
    End Function

    Public Function itemsPerSupplier(ByVal orderIDParam As String, ByVal deliveryIDParam As String) As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("SELECT distinct 
                                    quotItemListID,
                                    delivery.packingBuying,
									quotitemslist.quotItemPrice,
									quotitemslist.quotItemIncrement,
									deliveryitems.deliveryDeliveredQty,
                                    quotitemslist.quotItemPrice,
                                    supplierquotations.supplierDiscountPer,
                                    supplierquotations.supplierDiscountLS,  
                                    supplierquotations.supplierCurrencyCode

                                    FROM easyenquiry.delivery
                                    LEFT JOIN easyenquiry.orders ON delivery.order_ID = orders.orderID
									LEFT JOIN easyenquiry.quotations ON orders.quotation_ID = quotations.quotationID                                   
                                    LEFT JOIN easyenquiry.quotitemslist ON quotations.quotationID = quotitemslist.quotation_ID
									LEFT JOIN easyenquiry.deliveryitems ON quotitemslist.quotItemListID = deliveryitems.quotItemList_ID
                                    LEFT JOIN easyenquiry.suppliersinvoices ON delivery.deliveryID = suppliersinvoices.delivery_ID
									LEFT JOIN easyenquiry.supplierquotations ON delivery.supQuotation_ID = supplierquotations.supQuotationID

                                    WHERE orderID = '" & orderIDParam & "'  AND deliveryID = '" & deliveryIDParam & "'
                                    and quotItemPrice<>0 AND deliveryDeliveredQty <> 0
									ORDER BY quotItemListID ASC ", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(finalOrderSummary, "finalordersum")
        Return finalOrderSummary
    End Function

    Public Function deliveryPerOrder(ByVal orderIDParam As String) As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("SELECT DISTINCT deliveryID
									FROM easyenquiry.easyenquiry.delivery
									Where order_ID = '" & orderIDParam & "'", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(finalOrderSummary, "finalordersum")
        Return finalOrderSummary
    End Function

    Public Function extraChargesPerDelivery(ByVal deliveryIDParam As String) As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("SELECT DISTINCT freightBuying, freightBuyingCurr, customsBuying, customsBuyingCurr, extraCharges1Buying, extraCharges1BuyingCurr, extraCharges2Buying, extraCharges2BuyingCurr, extraCharges3Buying, extraCharges3BuyingCurr
									FROM easyenquiry.easyenquiry.delivery
									WHERE deliveryID =  '" & deliveryIDParam & "'", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(finalOrderSummary, "finalordersum")
        Return finalOrderSummary
    End Function

    Public Function invoiceForCNBroker(ByVal deliveryIDParam As String) As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("Select invoiceNumber From easyenquiry.easyenquiry.invoices
									WHERE delivery_ID =  '" & deliveryIDParam & "'", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(finalOrderSummary, "finalordersum")
        Return finalOrderSummary
    End Function

    Public Function deliveryBrokerCode(ByVal creditNoteNumber As String) As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("Select deliveryBroker_ID From easyenquiry.easyenquiry.creditnotes
									WHERE creditNoteNumber =  '" & creditNoteNumber & "'", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(finalOrderSummary, "finalordersum")
        Return finalOrderSummary
    End Function

    Public Function exportedInvCN(ByVal deliveryIDParam As String) As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("SELECT invoices.exported as invExported,
                                    creditnotes.exported as cnExported
                                    From easyenquiry.easyenquiry.delivery
                                    Left Join easyenquiry.easyenquiry.invoices On deliveryID = invoices.delivery_ID
                                    Left Join easyenquiry.easyenquiry.creditnotes ON deliveryID = creditnotes.delivery_ID
									WHERE deliveryID =  '" & deliveryIDParam & "'", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(finalOrderSummary, "finalordersum")
        Return finalOrderSummary
    End Function

    Public Function aspList() As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("SELECT adminusers.userFullName,
                    adminusers.userPrefix
                    FROM easyenquiry.easyenquiry.adminusers
                    WHERE userJobPosition like '%After Sales Person%'", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(finalOrderSummary, "finalordersum")
        Return finalOrderSummary
    End Function

    Public Function supQuotItems(ByVal supQuotationID As String) As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter("SELECT supQuotItemListID FROM easyenquiry.easyenquiry.supquotitemslist WHERE supQuotation_ID = '" & supQuotationID & "'", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(finalOrderSummary, "finalordersum")
        Return finalOrderSummary
    End Function

    Public Function loadItemsList(ByVal enquiryIDParam As String, Optional ByVal supQuotationIDParam As String = "") As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter

        If supQuotationIDParam = "" Then
            daFeatures.SelectCommand = New SqlCommand("SELECT * FROM easyenquiry.easyenquiry.itemslist WHERE enquiry_ID='" & enquiryIDParam & "'", cnDelivery)
        Else
            daFeatures.SelectCommand = New SqlCommand("SELECT supquotitemslist.supQuotItemListID, supquotitemslist.supQuotItemPartNumber AS itemPartNumber, supquotitemslist.supQuotItemDesc As itemDesc, supquotitemslist.supQuotItemQuantity As itemQuantity, supquotitemslist.supQuotItemUnit As itemUnit, supquotitemslist.supQuotItemDeliveryDays As itemDeliveryDays, supquotitemslist.supQuotItemPrice As itemPrice , supplierquotations.supplierDiscountPer, supplierquotations.supplierDiscountLS FROM easyenquiry.easyenquiry.supquotitemslist LEFT JOIN easyenquiry.easyenquiry.supplierquotations ON supquotitemslist.supQuotation_ID=supplierquotations.supQuotationID WHERE supplierquotations.enquiry_ID='" & enquiryIDParam & "' AND supquotitemslist.supQuotation_ID='" & supQuotationIDParam & "'", cnDelivery)
        End If

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(finalOrderSummary, "finalordersum")
        Return finalOrderSummary
    End Function

    Public Function loadItemsListCustQuot(ByVal enquiryIDParam As String, ByVal supEvalQuotIDParam As String, Optional ByVal quotationIDParam As String = "") As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter

        If (quotationIDParam <> "") Then
            daFeatures.SelectCommand = New SqlCommand("SELECT quotitemslist.quotItemListID, supquotitemslist.supQuotItemListID AS itemSupQuotItemListID, supquotitemslist.supQuotation_ID, supplierevalquotationlist.supEvalQuotListID As itemSupEvalQuotListID, suppliers.supplierID AS itemSupplierID, suppliers.supplierName AS itemSupplierName, supquotitemslist.supQuotItemPartNumber AS itemPartNumber, supquotitemslist.supQuotItemDesc AS itemDesc, quotitemslist.quotItemDeliveryDays AS itemDeliveryDays, quotitemslist.quotItemQuantity AS itemQuantity, supquotitemslist.supQuotItemUnit AS itemUnit, supplierquotations.supplierCurrencyCode AS itemSupplierCurrencyCode, quotitemslist.quotItemPrice  AS itemPrice, quotitemslist.quotItemIncrement AS itemIncrement, quotations.customerCurrencyCode, quotations.customerDiscountPer, quotations.customerDiscountLS,case when quotations.broker_ID IS NULL then ''else  quotations.broker_ID end AS broker_ID, quotations.brokerageFeePer, quotations.brokerageFeeLS,  quotations.quotationFreightDestination from easyenquiry.easyenquiry.supplierevalquotation left join easyenquiry.easyenquiry.supplierevalquotationlist ON supplierevalquotation.supEvalQuotID=supplierevalquotationlist.supEvalQuot_ID left join easyenquiry.easyenquiry.supplierquotations ON supplierevalquotationlist.supQuotation_ID=supplierquotations.supQuotationID left join easyenquiry.easyenquiry.supquotitemslist ON supplierquotations.supQuotationID=supquotitemslist.supQuotation_ID left join easyenquiry.easyenquiry.suppliers ON supplierquotations.supplier_ID=suppliers.supplierID left join easyenquiry.easyenquiry.quotations ON supplierevalquotation.supEvalQuotID=quotations.supEvalQuot_ID left join easyenquiry.easyenquiry.quotitemslist ON quotations.quotationID=quotitemslist.quotation_ID WHERE quotitemslist.supEvalQuotList_ID=supplierevalquotationlist.supEvalQuotListID AND supplierevalquotationlist.supQuotItemList_ID=supquotitemslist.supQuotItemListID AND supplierquotations.enquiry_ID='" & enquiryIDParam & "' AND supplierevalquotation.supEvalQuotID='" & supEvalQuotIDParam & "' AND quotations.quotationID='" & quotationIDParam & "' ORDER BY supplierevalquotationlist.supEvalQuotListPos ASC", cnDelivery)
        Else
            daFeatures.SelectCommand = New SqlCommand("SELECT supquotitemslist.supQuotItemListID AS itemSupQuotItemListID, supquotitemslist.supQuotation_ID, supplierevalquotationlist.supEvalQuotListID As itemSupEvalQuotListID ,suppliers.supplierID AS itemSupplierID, suppliers.supplierName AS itemSupplierName, supquotitemslist.supQuotItemPartNumber AS itemPartNumber, supquotitemslist.supQuotItemDesc AS itemDesc, supquotitemslist.supQuotItemDeliveryDays AS itemDeliveryDays, supquotitemslist.supQuotItemQuantity AS itemQuantity, supquotitemslist.supQuotItemUnit AS itemUnit, supplierquotations.supplierCurrencyCode AS itemSupplierCurrencyCode, supquotitemslist.supQuotItemPrice  AS itemPrice from easyenquiry.easyenquiry.supplierevalquotation left join easyenquiry.easyenquiry.supplierevalquotationlist ON supplierevalquotation.supEvalQuotID=supplierevalquotationlist.supEvalQuot_ID left join easyenquiry.easyenquiry.supplierquotations ON supplierevalquotationlist.supQuotation_ID=supplierquotations.supQuotationID left join easyenquiry.easyenquiry.supquotitemslist ON supplierquotations.supQuotationID=supquotitemslist.supQuotation_ID left join easyenquiry.easyenquiry.suppliers ON supplierquotations.supplier_ID=suppliers.supplierID WHERE supplierevalquotationlist.supQuotItemList_ID=supquotitemslist.supQuotItemListID AND supplierquotations.enquiry_ID='" & enquiryIDParam & "' AND supplierevalquotation.supEvalQuotID='" & supEvalQuotIDParam & "' ORDER BY supplierevalquotationlist.supEvalQuotListPos ASC", cnDelivery)
        End If

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(finalOrderSummary, "finalordersum")
        Return finalOrderSummary
    End Function

    Public Function loadItemsListDelItems(ByVal sqlCommand As String) As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter

        daFeatures.SelectCommand = New SqlCommand(sqlCommand, cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(finalOrderSummary, "finalordersum")
        Return finalOrderSummary
    End Function

    Public Function loadSupDueDate(ByVal orderIDParam As String, ByVal mergedIDParam As String, Optional ByVal deliveryIDParam As String = "") As DataSet
        Dim finalOrderSummary = New DataSet

        Dim cnDelivery As SqlConnection = New SqlConnection(SQL_CONNECT)
        Dim daFeatures As New SqlDataAdapter

        daFeatures.SelectCommand = New SqlCommand("SELECT 
                                case when ((SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =supplierquotations.supplierPaymentTerm_ID) > 0) 
                                then 
                                case when convert(date, CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =supplierquotations.supplierPaymentTerm_ID) , 
                                case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =supplierquotations.supplierPaymentTerm_ID) = 'dispatchDate' 
                                then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =supplierquotations.supplierPaymentTerm_ID) = 'invoiceDate' 
                                then  (select suppliersinvoices.supplierInvoiceDate from easyenquiry.easyenquiry.suppliersinvoices where delivery.deliveryID = suppliersinvoices.delivery_ID) else deliveryDate end )) AS DATE),103),103)
                                < convert(date, CONVERT(VARCHAR(10),CAST((select suppliersinvoices.supplierInvoiceDate from easyenquiry.easyenquiry.suppliersinvoices where delivery.deliveryID = suppliersinvoices.delivery_ID) AS DATE),103),103) 
                                then CONVERT(VARCHAR(10),CAST(invoiceDate AS DATE),103) else CONVERT(VARCHAR(10),CAST( (SELECT dateadd (day,(SELECT paymentTermValueInDays from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =supplierquotations.supplierPaymentTerm_ID) , 
                                case when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =supplierquotations.supplierPaymentTerm_ID) = 'dispatchDate' 
                                then dispatchDate when (SELECT paymentTermCountFrom from easyenquiry.easyenquiry.paymentterms WHERE paymentterms.paymentTermID =supplierquotations.supplierPaymentTerm_ID) = 'invoiceDate' 
                                then  (select suppliersinvoices.supplierInvoiceDate from easyenquiry.easyenquiry.suppliersinvoices where delivery.deliveryID = suppliersinvoices.delivery_ID) else deliveryDate end)) AS DATE),103) end else CONVERT(VARCHAR(10),CAST((select suppliersinvoices.supplierInvoiceDate from easyenquiry.easyenquiry.suppliersinvoices where delivery.deliveryID = suppliersinvoices.delivery_ID) AS DATE),103) end AS supplierDueDate2 
                                FROM easyenquiry.easyenquiry.invoices LEFT JOIN easyenquiry.easyenquiry.deliverymerged ON invoices.deliveryMerged_ID=deliverymerged.deliveryMergedID LEFT JOIN easyenquiry.easyenquiry.delivery on deliverymerged.delivery_ID= delivery.deliveryID LEFT JOIN easyenquiry.easyenquiry.suppliersinvoices on suppliersinvoices.delivery_ID=delivery.deliveryID LEFT JOIN easyenquiry.easyenquiry.orders on delivery.order_ID=orders.orderID LEFT JOIN easyenquiry.easyenquiry.quotations ON orders.quotation_ID= quotations.quotationID  LEFT JOIN easyenquiry.easyenquiry.supplierquotations ON supplierquotations.supQuotationID= delivery.supQuotation_ID LEFT JOIN easyenquiry.easyenquiry.suppliers ON suppliers.supplierID =supplierquotations.supplier_ID LEFT JOIN easyenquiry.easyenquiry.paymentterms ON paymentterms.paymentTermID=quotations.paymentTerm_ID WHERE orderID='" & orderIDParam & "' AND deliveryMerged_ID = '" & mergedIDParam & "' and deliveryID = '" & deliveryIDParam & "'", cnDelivery)

        daFeatures.SelectCommand.CommandTimeout = 0

        ' data adapter filling dataset
        daFeatures.Fill(finalOrderSummary, "finalordersum")
        Return finalOrderSummary
    End Function

End Class