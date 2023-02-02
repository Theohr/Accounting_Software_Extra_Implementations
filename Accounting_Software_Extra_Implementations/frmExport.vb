Imports System.IO
Imports System.Data.SqlClient
Imports System.Data.OracleClient

Public Class frmExport

    ' Public Variables
    Dim orderID As Integer
    Dim generaldataset As New datasetsAndTables()
    Dim sellingArray(,) As String
    Dim buyingArray(,) As String
    Dim creditNoteID As String = ""
    Dim gen = New ExportDanaosData()
    Dim debitNoteID As String = ""
    Dim adjustmentID As String = ""

    Public Sub New()
        InitializeComponent()
    End Sub
    Public Sub setCurrentCN(ByVal currentCnParam As String)
        creditNoteID = currentCnParam
    End Sub

    Public Sub setCurrentDN(ByVal currentDnParam As String)
        debitNoteID = currentDnParam
    End Sub
    Public Sub setCurrentAN(ByVal currentAnParam As String)
        adjustmentID = currentAnParam
    End Sub
    Public Sub setOrderID(ByVal orderIDParam As String)
        orderID = orderIDParam
    End Sub

    Public Sub exportData()

        danaosUsername = danaosUsername.Substring(0, 4)

        ' Creating the dataset in loadQUery function in datasetsandTables class
        Dim finalized = generaldataset.finalizedCheck(orderID)
        Dim exported = generaldataset.exportedCheck(orderID)
        Dim cnExported = generaldataset.cnExportedCheck(orderID)
        Dim dnExported = generaldataset.dnExportedCheck(orderID)
        Dim anExported = generaldataset.anExportedCheck(orderID)
        Dim dates = generaldataset.getDates(orderID, creditNoteID)
        Dim supinvoices = generaldataset.getSupplierInvoices(orderID)

        Dim finBool = finalizedBool(finalized)
        Dim expBool = exportedBool(exported)
        Dim cnExpBool = cnBool(cnExported)
        Dim dnExpBool = dnBool(dnExported)
        Dim anExpBool = anBool(anExported)
        Dim datesBool = datesMatch(dates)
        Dim supInvBool = supplierInvoices(supinvoices)
        Dim creditNoteBroker
        Try
            creditNoteBroker = dates.Tables(0).Rows(0).Item("creditNoteType_ID")
        Catch ex As Exception
            creditNoteBroker = 1
        End Try


        If finBool = True And expBool = True Or cnExpBool = True Or dnExpBool = True Or anExpBool = True Then
            'Create a form object
            Dim oForm As New frmDataGrid
            ' Send CreditNoteID on select
            oForm.setCurrentCN(creditNoteID)
            If userName <> "christina.c" And userName <> "christina.k" And userName <> "vasia.c" And userName <> "antonia.t" Then
                gen.setCurrentCN(creditNoteID)
                gen.setCurrentDN(debitNoteID)
                gen.setCurrentAN(adjustmentID)

                oForm.Show()
            Else
                gen.setCurrentCN(creditNoteID)
                gen.setCurrentDN(debitNoteID)
                gen.setCurrentAN(adjustmentID)

                If datesBool = True Or creditNoteBroker <> 2 Or anExpBool = True Then
                    If supInvBool = True Then
                        gen.DanaosData()
                    Else
                        MessageBox.Show("Supplier Invoice is missing from one of the deliveries you are trying to export.", "Important Message", MessageBoxButtons.OK)
                    End If
                Else
                    MessageBox.Show("Journal Dates of Invoice and Credit Note does not match.", "Important Message", MessageBoxButtons.OK)
                End If
            End If
        ElseIf finBool = False Or expBool = True Then
            MessageBox.Show("Invoices are not finalized yet.", "Important Message", MessageBoxButtons.OK)
        ElseIf finBool = True Or expBool = False Then
            MessageBox.Show("Invoices are already exported.", "Important Message", MessageBoxButtons.OK)
        End If
        'calls the datagrid window
    End Sub

    Public Function calculateSelling() As String(,)
        '======================================================================================================================'
        'SELLING ITEMS
        ' Dim deliveryIDarray(rowsCount) As Integer
        Dim dsSellingInvoices = generaldataset.loadSellingInvoices(orderID)
        Dim rowsCount = dsSellingInvoices.Tables("invoicesSelling").Rows.Count - 1
        Dim invRowsCount = (rowsCount + 1) * 3
        Dim invoicesSellArray(invRowsCount, 24) As String
        Dim nextIndex As Integer = 0
        Dim mergedID As Integer = 0
        Dim invNum As String = ""
        Dim dtSelling As DataSet
        Dim dtExchange As DataSet
        Dim dtPacking As DataSet
        Dim j As Integer = 0
        'Dim dtLoading = generaldataset.loadQuery(orderID)
        Dim rate = 0.0
        Dim exchanged = 0.0
        ' For loop for distinct selling invoicenumbers/mergedIDs
        For i = 0 To rowsCount

            'FindLastPk(dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("deliveryMergedID"))

            If dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("exported") = 0 And dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("finalized") = 1 Then
                'Assing the ecported value of each id
                Dim sellingSum As Decimal = 0.0

                'Assign the first row values in the variables
                invNum = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("invoiceNumber")
                mergedID = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("deliveryMerged_ID")

                'Get the selling dataset
                dtSelling = generaldataset.loadItemsSelling(orderID, mergedID, invNum)
                'Get Exchange Rates
                dtExchange = generaldataset.loadExchangeRates(orderID)
                ' Get packing
                dtPacking = generaldataset.loadPackingSelling(orderID, mergedID, invNum)

                If dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("invoiceType_ID") <> 3 Then
                    ' Go into a loop to sum the item prices w/Increment
                    For w = 0 To dtSelling.Tables("itemsSelling").Rows.Count - 1

                        For h = 0 To dtExchange.Tables(0).Rows.Count - 1
                            Try
                                If dtSelling.Tables(0).Rows(w).Item("supplierInvoiceNumber") = dtExchange.Tables(0).Rows(h).Item("supplierInvoiceNumber") Or dtSelling.Tables(0).Rows(w).Item("supplierInvoiceNumber") = "0" Then
                                    rate = dtExchange.Tables(0).Rows(h).Item("supplierExchangeRate")
                                    exchanged = dtSelling.Tables("itemsSelling").Rows(w).Item("quotItemPrice") / rate * dsSellingInvoices.Tables(0).Rows(i).Item("customerExchangeRate")
                                    'exchanged = Math.Round(dtSelling.Tables("itemsSelling").Rows(w).Item("quotItemPrice") / rate * dsSellingInvoices.Tables(0).Rows(i).Item("customerExchangeRate"), 2, MidpointRounding.AwayFromZero)
                                    ' Add increment
                                    'Dim multiQtyPr = dtSelling.Tables("itemsSelling").Rows(w).Item("quotItemPrice") + (dtSelling.Tables("itemsSelling").Rows(w).Item("quotItemPrice") * dtSelling.Tables("itemsSelling").Rows(w).Item("quotItemIncrement") / 100)
                                    Dim multiQtyPr = exchanged + (exchanged * dtSelling.Tables("itemsSelling").Rows(w).Item("quotItemIncrement") / 100)

                                    ' Get the regex expression
                                    Dim regex = New System.Text.RegularExpressions.Regex("(?<=[\\.])[0-9]+")

                                    ' If number matches regex
                                    If (regex.IsMatch(multiQtyPr)) Then
                                        ' Get the number after the dot
                                        Dim decimal_places As String = regex.Match(multiQtyPr).Value
                                        ' if the number exceeds the 2 digits then dont do anything
                                        If decimal_places.Length >= 2 Then
                                            Dim number As String = ""
                                            Try
                                                number = decimal_places.ToString()(2)
                                            Catch ex As Exception
                                                GoTo roundingToEven
                                            End Try
                                            Dim charToInt As Integer = Convert.ToInt32(Number)
                                            If charToInt >= 5 And charToInt < 10 Then
                                                multiQtyPr = multiQtyPr + 0.001
                                                multiQtyPr = Math.Round(multiQtyPr, 2, MidpointRounding.AwayFromZero)
                                            Else
roundingToEven:
                                                multiQtyPr = Math.Round(multiQtyPr, 2, MidpointRounding.ToEven)
                                            End If
                                            'If number = "5" Or number = "6" Or number = "7" Or number = "9" Or number = "8" Then
                                            '    multiQtyPr = Math.Round(multiQtyPr, 2, MidpointRounding.AwayFromZero)
                                            'Else
                                            '    multiQtyPr = Math.Round(multiQtyPr, 2, MidpointRounding.ToEven)
                                            'End If
                                        End If
                                    End If
                                    ' Multiply quantity 
                                    Dim idk = Math.Round(dtSelling.Tables("itemsSelling").Rows(w).Item("deliveryDeliveredQty") * multiQtyPr, 2)
                                    ' Add to selling Sum
                                    sellingSum = sellingSum + idk
                                End If
                            Catch ex As Exception
                                MessageBox.Show("Exchange Rates not found for the delivery you are trying to export.", "Important Message", MessageBoxButtons.OK)
                            End Try
                        Next
                    Next
                End If

                If sellingSum > 0 Then
                    For v = 0 To dtSelling.Tables("itemsSelling").Rows.Count - 1
                        ' Applying the Customers Discount in the Sum Percentage Or Fixed Price on first row percentage or price
                        If dtSelling.Tables("itemsSelling").Rows(v).Item("customerDiscountPer") <> 0.0 Then
                            Dim idk2 = Math.Round((sellingSum * dtSelling.Tables("itemsSelling").Rows(v).Item("customerDiscountPer") / 100), 2, MidpointRounding.AwayFromZero)

                            sellingSum = sellingSum - Math.Round(idk2, 2)
                            GoTo endDiscountSelling
                        ElseIf dtSelling.Tables("itemsSelling").Rows(v).Item("customerDiscountLS") <> 0.0 Then
                            sellingSum = sellingSum - dtSelling.Tables("itemsSelling").Rows(v).Item("customerDiscountLS")
                            GoTo endDiscountSelling
                        End If
                    Next

endDiscountSelling:

                    ' Add packing
                    For h = 0 To dtPacking.Tables(0).Rows.Count - 1
                        If dtPacking.Tables(0).Rows(h).Item("packingSelling") <> 0 Then
                            sellingSum = sellingSum + dtPacking.Tables(0).Rows(h).Item("packingSelling")
                            GoTo endPackingS
                        End If
                    Next
                End If
endPackingS:


                ' loads selling charges freight etc
                Dim dtChargesSell = generaldataset.loadChargesSelling(orderID, mergedID, invNum)

                'Extra Charges, freight and customs vars
                Dim freightSelling = 0.0
                Dim customsSelling = 0.0
                Dim extraCharges1Selling = 0.0
                Dim extraCharges2Selling = 0.0
                Dim extraCharges3Selling = 0.0

                Dim sellVat = 0.0
                Dim freightSellVat = 0.0
                Dim customsSellVat = 0.0
                Dim extraCharges1SellVat = 0.0
                Dim extraCharges2SellVat = 0.0
                Dim extraCharges3SellVat = 0.0

                ' loop to get the extra charges, freight customs etc and the VAT for each invoiceNum/MergedID
                For e = 0 To dtChargesSell.Tables("extraCharges").Rows.Count - 1
                    freightSelling = freightSelling + dtChargesSell.Tables("extraCharges").Rows(e).Item("freightSelling")
                    customsSelling = customsSelling + dtChargesSell.Tables("extraCharges").Rows(e).Item("customsSelling")
                    extraCharges1Selling = extraCharges1Selling + dtChargesSell.Tables("extraCharges").Rows(e).Item("extraCharges1Selling")
                    extraCharges2Selling = extraCharges2Selling + dtChargesSell.Tables("extraCharges").Rows(e).Item("extraCharges2Selling")
                    extraCharges3Selling = extraCharges3Selling + dtChargesSell.Tables("extraCharges").Rows(e).Item("extraCharges3Selling")

                    freightSellVat = freightSellVat + calculateVat(freightSelling, dtChargesSell.Tables("extraCharges").Rows(e).Item("freightSellingVat_ID"))
                    customsSellVat = customsSellVat + calculateVat(customsSelling, dtChargesSell.Tables("extraCharges").Rows(e).Item("customsSellingVat_ID"))
                    extraCharges1SellVat = extraCharges1SellVat + calculateVat(extraCharges1Selling, dtChargesSell.Tables("extraCharges").Rows(e).Item("extraCharges1SellingVat_ID"))
                    extraCharges2SellVat = extraCharges2SellVat + calculateVat(extraCharges2Selling, dtChargesSell.Tables("extraCharges").Rows(e).Item("extraCharges2SellingVat_ID"))
                    extraCharges3SellVat = extraCharges3SellVat + calculateVat(extraCharges3Selling, dtChargesSell.Tables("extraCharges").Rows(e).Item("extraCharges3SellingVat_ID"))
                Next

                'Getting the VAT Amount for each one
                sellVat = sellVat + calculateVat(sellingSum, dsSellingInvoices.Tables(0).Rows(i).Item("customerVat_ID"))

                ' Adds the VAT Buying
                Dim vatTotalSell As Double = 0.0
                vatTotalSell = vatTotalSell + sellVat

                Dim dtData = generaldataset.loadOutputValues(orderID, mergedID)

                Dim currency = dtData.Tables("outputvalues").Rows(0).Item("customerCurrencyCode")

                If dtChargesSell.Tables(0).Rows.Count > 0 Then
                    If dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("invoiceType_ID") = 1 Then
                        sellingSum = sellingSum + freightSelling + customsSelling + extraCharges1Selling + extraCharges2Selling + extraCharges3Selling
                        vatTotalSell = vatTotalSell + freightSellVat + customsSellVat + extraCharges1SellVat + extraCharges2SellVat + extraCharges3SellVat
                    ElseIf dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("invoiceType_ID") = 2 Then
                        If freightSelling = 0.00 Then
                            sellingSum = sellingSum + customsSelling + freightSelling + extraCharges2Selling + extraCharges3Selling
                            vatTotalSell = vatTotalSell + customsSellVat + freightSellVat + extraCharges2SellVat + extraCharges3SellVat
                        Else
                            sellingSum = sellingSum + customsSelling + extraCharges1Selling + extraCharges2Selling + extraCharges3Selling
                            vatTotalSell = vatTotalSell + customsSellVat + extraCharges1SellVat + extraCharges2SellVat + extraCharges3SellVat
                        End If
                    ElseIf dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("invoiceType_ID") = 3 Then
                        If freightSelling = 0.00 Then
                            sellingSum = sellingSum + extraCharges1Selling
                            vatTotalSell = vatTotalSell + extraCharges1SellVat
                        Else
                            sellingSum = sellingSum + freightSelling
                            vatTotalSell = vatTotalSell + freightSellVat
                        End If
                        currency = dtChargesSell.Tables(0).Rows(0).Item("freightSellingCurr")
                    End If
                End If

                ' OLD CALCULATIONS
                'sellingSum = sellingSum + freightSelling + customsSelling + extraCharges1Selling + extraCharges2Selling + extraCharges3Selling          
                'vatTotalSell = vatTotalSell + sellVat + freightSellVat + customsSellVat + extraCharges1SellVat + extraCharges2SellVat + extraCharges3SellVat

                'Total Amount summed up
                Dim grandTotalSell As Double = 0.0
                grandTotalSell = grandTotalSell + sellingSum + vatTotalSell

                ' Get the rest of the data in a dataset of the order and merged id

                If dtData.Tables("outputvalues").Rows(0).Item("vesselName") = "NA" Then
                    dtData.Tables("outputvalues").Rows(0).Item("vesselName") = "NOT SPECIFIED"
                End If

                Dim narrative = dtData.Tables("outputvalues").Rows(0).Item("narativeOfEntry") + "-" + dtData.Tables("outputvalues").Rows(0).Item("vesselName") + " -PO " + dtData.Tables("outputvalues").Rows(0).Item("custPurchOrderRef")

                Dim ifExists = chkIfAllExist(dtData.Tables("outputvalues").Rows(0).Item("customerCode"))

                'sellingSum = Math.Round(sellingSum / (dsSellingInvoices.Tables(0).Rows(i).Item("supplierExchangeRate") / dsSellingInvoices.Tables(0).Rows(i).Item("customerExchangeRate")), 2, MidpointRounding.AwayFromZero)
                'grandTotalSell = Math.Round(grandTotalSell / (dsSellingInvoices.Tables(0).Rows(i).Item("supplierExchangeRate") / dsSellingInvoices.Tables(0).Rows(i).Item("customerExchangeRate")), 2, MidpointRounding.AwayFromZero)


                If grandTotalSell <> 0 Then
                    ' Check if vat is equal or greater than 0 to see if its debit or credit
                    If vatTotalSell > 0 Then

                        'j increment each time we insert a row in the array 
                        If j > 0 Then j += 1


                        'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                        invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("deliveryMerged_ID")
                        invoicesSellArray(j, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                        invoicesSellArray(j, 2) = dtData.Tables("outputvalues").Rows(0).Item("customerCode")
                        invoicesSellArray(j, 3) = dtData.Tables("outputvalues").Rows(0).Item("customerCode")
                        invoicesSellArray(j, 4) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("invoiceNumber")
                        invoicesSellArray(j, 5) = dtData.Tables("outputvalues").Rows(0).Item("invoiceDate")
                        invoicesSellArray(j, 6) = TruncString.Truncate(narrative.ToString(), 60)
                        invoicesSellArray(j, 7) = dtData.Tables("outputvalues").Rows(0).Item("dueDate")
                        invoicesSellArray(j, 8) = currency
                        invoicesSellArray(j, 9) = grandTotalSell.ToString()
                        invoicesSellArray(j, 10) = "0" 'VAT Number
                        invoicesSellArray(j, 11) = "7" 'VAT Category
                        invoicesSellArray(j, 12) = vatTotalSell.ToString()
                        invoicesSellArray(j, 13) = "1.7" 'INC_EXP_CATEGORY
                        invoicesSellArray(j, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                        invoicesSellArray(j, 15) = dtData.Tables("outputvalues").Rows(0).Item("invoiceDate")
                        invoicesSellArray(j, 16) = danaosUsername
                        invoicesSellArray(j, 17) = "D" 'DEBIT CREDI
                        invoicesSellArray(j, 18) = "" 'VESSEL_CODE
                        Try
                            invoicesSellArray(j, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                        Catch ex As Exception
                            invoicesSellArray(j, 19) = dtData.Tables("outputvalues").Rows(0).Item("refNum")
                        End Try
                        invoicesSellArray(j, 20) = "023" 'JOURNAL_COMPANY
                        invoicesSellArray(j, 21) = "40" 'JOURNALSERIES
                        invoicesSellArray(j, 22) = "0" ' BOOK VALUE
                        invoicesSellArray(j, 23) = "0" ' ENTRYRATE
                        If ifExists = True Then
                            invoicesSellArray(j, 24) = "True"
                        Else
                            invoicesSellArray(j, 24) = "False"
                        End If

                        ' Adds a Next Row of the Amount + VAT
                        j += 1
                        'Inserts the First Row of the Amount without VAT
                        'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                        invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("deliveryMerged_ID")
                        invoicesSellArray(j, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                        invoicesSellArray(j, 2) = dtData.Tables("outputvalues").Rows(0).Item("salesCode")
                        invoicesSellArray(j, 3) = dtData.Tables("outputvalues").Rows(0).Item("customerCode")
                        invoicesSellArray(j, 4) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("invoiceNumber")
                        invoicesSellArray(j, 5) = dtData.Tables("outputvalues").Rows(0).Item("invoiceDate")
                        invoicesSellArray(j, 6) = TruncString.Truncate(narrative.ToString(), 60)
                        invoicesSellArray(j, 7) = dtData.Tables("outputvalues").Rows(0).Item("dueDate")
                        invoicesSellArray(j, 8) = currency
                        invoicesSellArray(j, 9) = sellingSum.ToString()
                        invoicesSellArray(j, 10) = "0" 'VAT Number
                        invoicesSellArray(j, 11) = "2" 'VAT Category
                        invoicesSellArray(j, 12) = "0"
                        invoicesSellArray(j, 13) = "" 'INC_EXP_CATEGORY
                        invoicesSellArray(j, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                        invoicesSellArray(j, 15) = dtData.Tables("outputvalues").Rows(0).Item("invoiceDate")
                        invoicesSellArray(j, 16) = danaosUsername
                        invoicesSellArray(j, 17) = "C" 'DEBIT CREDI
                        invoicesSellArray(j, 18) = "" 'VESSEL_CODE
                        Try
                            invoicesSellArray(j, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                        Catch ex As Exception
                            invoicesSellArray(j, 19) = dtData.Tables("outputvalues").Rows(0).Item("refNum")
                        End Try
                        invoicesSellArray(j, 20) = "023" 'JOURNAL_COMPANY
                        invoicesSellArray(j, 21) = "40" 'JOURNALSERIES
                        invoicesSellArray(j, 22) = "0" ' BOOK VALUE
                        invoicesSellArray(j, 23) = "0" ' ENTRYRATE
                        If ifExists = True Then
                            invoicesSellArray(j, 24) = "True"
                        Else
                            invoicesSellArray(j, 24) = "False"
                        End If


                        'Adds a THird Row of the VAT AMOunt only
                        j += 1
                        'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                        invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("deliveryMerged_ID")
                        invoicesSellArray(j, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                        invoicesSellArray(j, 2) = "38602"
                        invoicesSellArray(j, 3) = dtData.Tables("outputvalues").Rows(0).Item("customerCode")
                        invoicesSellArray(j, 4) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("invoiceNumber")
                        invoicesSellArray(j, 5) = dtData.Tables("outputvalues").Rows(0).Item("invoiceDate")
                        invoicesSellArray(j, 6) = TruncString.Truncate(narrative.ToString(), 60)
                        invoicesSellArray(j, 7) = dtData.Tables("outputvalues").Rows(0).Item("dueDate")
                        invoicesSellArray(j, 8) = currency
                        invoicesSellArray(j, 9) = vatTotalSell.ToString()
                        invoicesSellArray(j, 10) = "0" 'VAT Number
                        invoicesSellArray(j, 11) = "2" 'VAT Category
                        invoicesSellArray(j, 12) = "0"
                        invoicesSellArray(j, 13) = "" 'INC_EXP_CATEGORY
                        invoicesSellArray(j, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                        invoicesSellArray(j, 15) = dtData.Tables("outputvalues").Rows(0).Item("invoiceDate")
                        invoicesSellArray(j, 16) = danaosUsername
                        invoicesSellArray(j, 17) = "C" 'DEBIT CREDI
                        invoicesSellArray(j, 18) = "" 'VESSEL_CODE
                        Try
                            invoicesSellArray(j, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                        Catch ex As Exception
                            invoicesSellArray(j, 19) = dtData.Tables("outputvalues").Rows(0).Item("refNum")
                        End Try
                        invoicesSellArray(j, 20) = "023" 'JOURNAL_COMPANY
                        invoicesSellArray(j, 21) = "40" 'JOURNALSERIES
                        invoicesSellArray(j, 22) = "0" ' BOOK VALUE
                        invoicesSellArray(j, 23) = "0" ' ENTRYRATE
                        If ifExists = True Then
                            invoicesSellArray(j, 24) = "True"
                        Else
                            invoicesSellArray(j, 24) = "False"
                        End If

                    Else
                        ' If VAT = 0 Then we do the same as above without the VAT ROW
                        If j > 0 Then j += 1

                        'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                        invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("deliveryMerged_ID")
                        invoicesSellArray(j, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                        invoicesSellArray(j, 2) = dtData.Tables("outputvalues").Rows(0).Item("customerCode")
                        invoicesSellArray(j, 3) = dtData.Tables("outputvalues").Rows(0).Item("customerCode")
                        invoicesSellArray(j, 4) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("invoiceNumber")
                        invoicesSellArray(j, 5) = dtData.Tables("outputvalues").Rows(0).Item("invoiceDate")
                        invoicesSellArray(j, 6) = TruncString.Truncate(narrative.ToString(), 60)
                        invoicesSellArray(j, 7) = dtData.Tables("outputvalues").Rows(0).Item("dueDate")
                        invoicesSellArray(j, 8) = currency
                        invoicesSellArray(j, 9) = grandTotalSell.ToString()
                        invoicesSellArray(j, 10) = "0" 'VAT Number
                        invoicesSellArray(j, 11) = "2" 'VAT Category
                        invoicesSellArray(j, 12) = vatTotalSell.ToString()
                        invoicesSellArray(j, 13) = "1.1" 'INC_EXP_CATEGORY
                        invoicesSellArray(j, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                        invoicesSellArray(j, 15) = dtData.Tables("outputvalues").Rows(0).Item("invoiceDate")
                        invoicesSellArray(j, 16) = danaosUsername
                        invoicesSellArray(j, 17) = "D" 'DEBIT CREDI
                        invoicesSellArray(j, 18) = "" 'VESSEL_CODE
                        Try
                            invoicesSellArray(j, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                        Catch ex As Exception
                            invoicesSellArray(j, 19) = dtData.Tables("outputvalues").Rows(0).Item("refNum")
                        End Try
                        invoicesSellArray(j, 20) = "023" 'JOURNAL_COMPANY
                        invoicesSellArray(j, 21) = "40" 'JOURNALSERIES
                        invoicesSellArray(j, 22) = "0" ' BOOK VALUE
                        invoicesSellArray(j, 23) = "0" ' ENTRYRATE
                        If ifExists = True Then
                            invoicesSellArray(j, 24) = "True"
                        Else
                            invoicesSellArray(j, 24) = "False"
                        End If

                        ' Row of Total with VAT
                        j += 1
                        ' Row of Total without VAT
                        'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                        invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("deliveryMerged_ID")
                        invoicesSellArray(j, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                        invoicesSellArray(j, 2) = dtData.Tables("outputvalues").Rows(0).Item("salesCode")
                        invoicesSellArray(j, 3) = dtData.Tables("outputvalues").Rows(0).Item("customerCode")
                        invoicesSellArray(j, 4) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("invoiceNumber")
                        invoicesSellArray(j, 5) = dtData.Tables("outputvalues").Rows(0).Item("invoiceDate")
                        invoicesSellArray(j, 6) = TruncString.Truncate(narrative.ToString(), 60)
                        invoicesSellArray(j, 7) = dtData.Tables("outputvalues").Rows(0).Item("dueDate")
                        invoicesSellArray(j, 8) = currency
                        invoicesSellArray(j, 9) = sellingSum.ToString()
                        invoicesSellArray(j, 10) = "0" 'VAT Number
                        invoicesSellArray(j, 11) = "2" 'VAT Category
                        invoicesSellArray(j, 12) = vatTotalSell.ToString()
                        invoicesSellArray(j, 13) = "" 'INC_EXP_CATEGORY
                        invoicesSellArray(j, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                        invoicesSellArray(j, 15) = dtData.Tables("outputvalues").Rows(0).Item("invoiceDate")
                        invoicesSellArray(j, 16) = danaosUsername
                        invoicesSellArray(j, 17) = "C" 'DEBIT CREDI
                        invoicesSellArray(j, 18) = "" 'VESSEL_CODE
                        Try
                            invoicesSellArray(j, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                        Catch ex As Exception
                            invoicesSellArray(j, 19) = dtData.Tables("outputvalues").Rows(0).Item("refNum")
                        End Try
                        invoicesSellArray(j, 20) = "023" 'JOURNAL_COMPANY
                        invoicesSellArray(j, 21) = "40" 'JOURNALSERIES
                        invoicesSellArray(j, 22) = "0" ' BOOK VALUE
                        invoicesSellArray(j, 23) = "0" ' ENTRYRATE
                        If ifExists = True Then
                            invoicesSellArray(j, 24) = "True"
                        Else
                            invoicesSellArray(j, 24) = "False"
                        End If
                    End If
                End If
            End If
        Next

        Return invoicesSellArray
        'END SELLING
        '====================================================================================================================='
    End Function

    Public Function calculateBuying() As String(,)
        '====================================================================================================================='
        'BUYING ITEMS
        ' Get the distinct values of buying dataset
        Dim dsBuyingInvoices = generaldataset.loadBuyingInvociesNew(orderID)
        Dim distRows = dsBuyingInvoices.Tables("invoicesBuying").Rows.Count - 1
        Dim iDistRows = (distRows + 1) * 3
        Dim invoicesBuyArray(iDistRows, 24) As String
        Dim dtBuying As DataSet
        Dim mergedIDuni As Integer
        Dim supInvNum As String
        Dim deliveryID As Integer
        Dim u As Integer
        Dim g = 0

        ' Run into a loop of the distinct values
        For y = 0 To distRows
            'chkIfAllExist(dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("InvoiceNumbers"))

            If dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("finalized") = 1 And dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("exported") = 0 Then
                Dim buyingSum As Double = 0.0

                Dim enqProcessStatusID = 0
                Dim delProcessStatusID = 0

                'Assign the first row values in the variables
                mergedIDuni = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("deliveryMergedID")
                supInvNum = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("InvoiceNumbers")
                deliveryID = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("deliveryID")


                'Get the buying dataset
                Try
                    If supInvNum.Contains("????INV") Then
                        supInvNum = supInvNum.Substring(0, supInvNum.IndexOf("-"))
                    End If
                Catch ex As Exception

                End Try

                dtBuying = generaldataset.loadItemsBuying(orderID, mergedIDuni, supInvNum)

                ' Go into a loop to sum the item prices
                For w = 0 To dtBuying.Tables("itemsBuying").Rows.Count - 1
                    Dim multyQTYPR = (dtBuying.Tables("itemsBuying").Rows(w).Item("deliveryReadyQty") * dtBuying.Tables("ItemsBuying").Rows(w).Item("quotItemPrice"))
                    buyingSum = buyingSum + Math.Round(multyQTYPR, 2, MidpointRounding.AwayFromZero)
                Next

                ' checks if buying sum greated than 0 to applu te discount and add the packing
                If buyingSum > 0 Then
                    For h = 0 To dtBuying.Tables("itemsBuying").Rows.Count - 1
                        ' Applying the Suppliers Discount in the Sum Percentage Or Fixed Price
                        If dtBuying.Tables("itemsBuying").Rows(h).Item("supplierDiscountPer") <> 0.0 Then
                            Dim idk = (buyingSum * dtBuying.Tables("itemsBuying").Rows(h).Item("supplierDiscountPer") / 100)
                            buyingSum = buyingSum - Math.Round(idk, 2, MidpointRounding.AwayFromZero)
                            GoTo endDiscount
                        ElseIf dtBuying.Tables("itemsBuying").Rows(h).Item("supplierDiscountLS") <> 0.0 Then
                            buyingSum = buyingSum - dtBuying.Tables("itemsBuying").Rows(h).Item("supplierDiscountLS")
                            GoTo endDiscount
                        End If
                    Next

endDiscount:

                    For h = 0 To dtBuying.Tables("itemsBuying").Rows.Count - 1
                        If dtBuying.Tables("itemsBuying").Rows(h).Item("packingBuying") <> 0 Then
                            buyingSum = buyingSum + dtBuying.Tables("itemsBuying").Rows(h).Item("packingBuying")
                            GoTo endPacking
                        End If
                    Next
                End If

endPacking:

                ' Receiving freight and other extra charges based on supplierNumber and Order ID
                'Dim dtCharges = generaldataset.loadCharges(orderID, deliveryID, supInvNum)
                Dim dtCharges = generaldataset.loadCharges1(orderID, mergedIDuni)

                'Extra Charges, freight and customs vars
                Dim freightBuying = 0.0
                Dim customsBuying = 0.0
                Dim extraCharges1Buying = 0.0
                Dim extraCharges2Buying = 0.0
                Dim extraCharges3Buying = 0.0
                Dim packingBuying = 0.0
                Dim buyVat = 0.0
                Dim freightBuyVat = 0.0
                Dim customsBuyVat = 0.0
                Dim extraCharges1BuyVat = 0.0
                Dim extraCharges2BuyVat = 0.0
                Dim extraCharges3BuyVat = 0.0

                'TODO: get the numbers in the array above and check if they are equal or not to sum the freight      
                'For i = 0 To dtCharges.Tables("extraCharges").Rows.Count - 1
                '    If deliveryID = dtCharges.Tables(0).Rows(i).Item("deliveryID") Or supInvNum = dtCharges.Tables(0).Rows(i).Item("supplierInvoiceNumber") Then
                '        freightBuying = freightBuying + dtCharges.Tables("extraCharges").Rows(i).Item("freightBuying")
                '        customsBuying = customsBuying + dtCharges.Tables("extraCharges").Rows(i).Item("customsBuying")
                '        extraCharges1Buying = extraCharges1Buying + dtCharges.Tables("extraCharges").Rows(i).Item("extraCharges1Buying")
                '        extraCharges2Buying = extraCharges2Buying + dtCharges.Tables("extraCharges").Rows(i).Item("extraCharges2Buying")
                '        extraCharges3Buying = extraCharges3Buying + dtCharges.Tables("extraCharges").Rows(i).Item("extraCharges3Buying")

                '        'Getting the VAT Amount for each one
                '        buyVat = buyVat + calculateVat(buyingSum, dtCharges.Tables("extraCharges").Rows(i).Item("supplierVat_ID"))
                '        freightBuyVat = freightBuyVat + calculateVat(freightBuying, dtCharges.Tables("extraCharges").Rows(i).Item("freightBuyingVat_ID"))
                '        customsBuyVat = customsBuyVat + calculateVat(customsBuying, dtCharges.Tables("extraCharges").Rows(i).Item("customsBuyingVat_ID"))
                '        extraCharges1BuyVat = extraCharges1BuyVat + calculateVat(extraCharges1Buying, dtCharges.Tables("extraCharges").Rows(i).Item("extraCharges1BuyingVat_ID"))
                '        extraCharges2BuyVat = extraCharges2BuyVat + calculateVat(extraCharges2Buying, dtCharges.Tables("extraCharges").Rows(i).Item("extraCharges2BuyingVat_ID"))
                '        extraCharges3BuyVat = extraCharges3BuyVat + calculateVat(extraCharges3Buying, dtCharges.Tables("extraCharges").Rows(i).Item("extraCharges3BuyingVat_ID"))
                '    End If
                'Next

                For i = 0 To dtCharges.Tables("extraCharges").Rows.Count - 1
                    If deliveryID = dtCharges.Tables(0).Rows(i).Item("deliveryID") Then
                        If supInvNum = dtCharges.Tables(0).Rows(i).Item("forwarderInvoiceNumber") Then
                            freightBuying = freightBuying + dtCharges.Tables("extraCharges").Rows(i).Item("freightBuying")
                            freightBuyVat = freightBuyVat + calculateVat(freightBuying, dtCharges.Tables("extraCharges").Rows(i).Item("freightBuyingVat_ID"))
                        End If
                        If supInvNum = dtCharges.Tables(0).Rows(i).Item("clearingAgentInvoiceNumber") Then
                            customsBuying = customsBuying + dtCharges.Tables("extraCharges").Rows(i).Item("customsBuying")
                            customsBuyVat = customsBuyVat + calculateVat(customsBuying, dtCharges.Tables("extraCharges").Rows(i).Item("customsBuyingVat_ID"))
                        End If
                        If supInvNum = dtCharges.Tables(0).Rows(i).Item("generalSupplier1InvoiceNumber") Then
                            extraCharges1Buying = extraCharges1Buying + dtCharges.Tables("extraCharges").Rows(i).Item("extraCharges1Buying")
                            extraCharges1BuyVat = extraCharges1BuyVat + calculateVat(extraCharges1Buying, dtCharges.Tables("extraCharges").Rows(i).Item("extraCharges1BuyingVat_ID"))
                        End If
                        If supInvNum = dtCharges.Tables(0).Rows(i).Item("generalSupplier2InvoiceNumber") Then
                            extraCharges2Buying = extraCharges2Buying + dtCharges.Tables("extraCharges").Rows(i).Item("extraCharges2Buying")
                            extraCharges2BuyVat = extraCharges2BuyVat + calculateVat(extraCharges2Buying, dtCharges.Tables("extraCharges").Rows(i).Item("extraCharges2BuyingVat_ID"))
                        End If
                        If supInvNum = dtCharges.Tables(0).Rows(i).Item("generalSupplier3InvoiceNumber") Then
                            extraCharges3Buying = extraCharges3Buying + dtCharges.Tables("extraCharges").Rows(i).Item("extraCharges3Buying")
                            extraCharges3BuyVat = extraCharges3BuyVat + calculateVat(extraCharges3Buying, dtCharges.Tables("extraCharges").Rows(i).Item("extraCharges3BuyingVat_ID"))
                        End If

                        'Getting the VAT Amount for each one
                    End If
                Next

                buyVat = buyVat + calculateVat(buyingSum, dsBuyingInvoices.Tables(0).Rows(y).Item("supplierVat_ID"))

                ' Adds the VAT Buying
                Dim vatTotal As Double = 0.0
                vatTotal = vatTotal + buyVat

                ' Get the rest of the data in a dataset of the order and merged id
                Dim dtData = generaldataset.loadOutputValues(orderID, mergedIDuni, deliveryID)
                Dim dtSupDueDate = generaldataset.loadSupDueDate(orderID, mergedIDuni, deliveryID)

                Dim enquiryID = dtData.Tables("outputvalues").Rows(0).Item("enquiryID")
                enqProcessStatusID = dtData.Tables("outputvalues").Rows(0).Item("processStatus_ID")

                Try

                    If enqProcessStatusID <> 6 Then
                        Dim SQL5 = "UPDATE easyenquiry.easyenquiry.enquiries SET processStatus_ID='6' WHERE enquiryID='" & enquiryID & "';"

                        Try
                            myCommand.Connection = conn
                            conn.Close()
                            conn.Open()
                            myCommand.CommandText = SQL5
                            myCommand.ExecuteNonQuery()
                            conn.Close()
                        Catch ex As Exception

                        End Try
                    End If
                Catch ex As Exception

                End Try

                Try
                    Dim SQL6 = "UPDATE easyenquiry.easyenquiry.delivery SET deliveryProcessStatus_ID='4' WHERE deliveryID='" & deliveryID & "';"

                    Try
                        myCommand.Connection = conn
                        conn.Close()
                        conn.Open()
                        myCommand.CommandText = SQL6
                        myCommand.ExecuteNonQuery()
                        conn.Close()
                    Catch ex As Exception

                    End Try
                Catch ex As Exception

                End Try

                Dim currency = ""

                Try
                    currency = dtBuying.Tables(0).Rows(0).Item("supplierCurrencyCode")
                Catch ex As Exception

                End Try

                'Try
                '    If dtData.Tables(0).Rows.Count > 1 And buyingSum > 0 Then
                '        currency = dtData.Tables("outputvalues").Rows(g).Item("supplierCurrencyCode")
                '        g = g + 1
                '    ElseIf dtData.Tables(0).Rows.Count = 1 Then
                '        currency = dtData.Tables("outputvalues").Rows(0).Item("supplierCurrencyCode")
                '    Else
                '        currency = dtData.Tables("outputvalues").Rows(0).Item("supplierCurrencyCode")
                '    End If
                'Catch
                '    MessageBox.Show("Currency not found.", "Important Message", MessageBoxButtons.OK)
                'End Try

                ' Extra Charges, Freight, Customs depending if its merged or split with each ones VAT
                If dtCharges.Tables("extraCharges").Rows.Count > 0 Then
                    For b = 0 To dtCharges.Tables(0).Rows.Count - 1
                        If deliveryID = dtCharges.Tables(0).Rows(b).Item("deliveryID") Then
                            ' Freight Check + VAT If different inv num then split else if merge
                            If dtCharges.Tables("extraCharges").Rows(b).Item("forwarderInvoiceNumber") <> "" And dtCharges.Tables("extraCharges").Rows(b).Item("forwarderInvoiceNumber") = supInvNum And dsBuyingInvoices.Tables(0).Rows(y).Item("Names") = "forwarderInvoiceNumber" Then
                                buyingSum = buyingSum + freightBuying
                                vatTotal = vatTotal + freightBuyVat
                                currency = dtCharges.Tables("extraCharges").Rows(0).Item("freightBuyingCurr")
                            ElseIf (dtCharges.Tables("extraCharges").Rows(0).Item("forwarderInvoiceNumber") = "" Or dtCharges.Tables("extraCharges").Rows(b).Item("forwarderInvoiceNumber") = supInvNum) And dsBuyingInvoices.Tables(0).Rows(y).Item("Names") = "supplierInvoiceNumber" Then
                                buyingSum = buyingSum + freightBuying
                                vatTotal = vatTotal + freightBuyVat
                            End If

                            'Customs + VAT If different inv num then split else if merge
                            If dtCharges.Tables("extraCharges").Rows(b).Item("clearingAgentInvoiceNumber") <> "" And dtCharges.Tables("extraCharges").Rows(b).Item("clearingAgentInvoiceNumber") = supInvNum And dsBuyingInvoices.Tables(0).Rows(y).Item("Names") = "clearingAgentInvoiceNumber" Then
                                buyingSum = buyingSum + customsBuying
                                vatTotal = vatTotal + customsBuyVat
                                currency = dtCharges.Tables("extraCharges").Rows(0).Item("customsBuyingCurr")
                            ElseIf (dtCharges.Tables("extraCharges").Rows(0).Item("clearingAgentInvoiceNumber") = "" Or dtCharges.Tables("extraCharges").Rows(b).Item("clearingAgentInvoiceNumber") = supInvNum) And dsBuyingInvoices.Tables(0).Rows(y).Item("Names") = "supplierInvoiceNumber" Then
                                buyingSum = buyingSum + customsBuying
                                vatTotal = vatTotal + customsBuyVat
                            End If

                            ' Extra Charges 1 Check + VAT If different inv num then split else if merge
                            If dtCharges.Tables("extraCharges").Rows(b).Item("generalSupplier1InvoiceNumber") <> "" And dtCharges.Tables("extraCharges").Rows(b).Item("generalSupplier1InvoiceNumber") = supInvNum And dsBuyingInvoices.Tables(0).Rows(y).Item("Names") = "generalSupplier1InvoiceNumber" Then
                                buyingSum = buyingSum + extraCharges1Buying
                                vatTotal = vatTotal + extraCharges1BuyVat
                                currency = dtCharges.Tables("extraCharges").Rows(0).Item("extraCharges1BuyingCurr")
                            ElseIf (dtCharges.Tables("extraCharges").Rows(0).Item("generalSupplier1InvoiceNumber") = "" Or dtCharges.Tables("extraCharges").Rows(b).Item("generalSupplier1InvoiceNumber") = supInvNum) And dsBuyingInvoices.Tables(0).Rows(y).Item("Names") = "supplierInvoiceNumber" Then
                                buyingSum = buyingSum + extraCharges1Buying
                                vatTotal = vatTotal + extraCharges1BuyVat
                            End If

                            'Extra Charges 2 Check = VAT If different inv num then split else if merge
                            If dtCharges.Tables("extraCharges").Rows(b).Item("generalSupplier2InvoiceNumber") <> "" And dtCharges.Tables("extraCharges").Rows(b).Item("generalSupplier2InvoiceNumber") = supInvNum And dsBuyingInvoices.Tables(0).Rows(y).Item("Names") = "generalSupplier2InvoiceNumber" Then
                                buyingSum = buyingSum + extraCharges2Buying
                                vatTotal = vatTotal + extraCharges2BuyVat
                                currency = dtCharges.Tables("extraCharges").Rows(0).Item("extraCharges2BuyingCurr")
                            ElseIf (dtCharges.Tables("extraCharges").Rows(0).Item("generalSupplier2InvoiceNumber") = "" Or dtCharges.Tables("extraCharges").Rows(b).Item("generalSupplier2InvoiceNumber") = supInvNum) And dsBuyingInvoices.Tables(0).Rows(y).Item("Names") = "supplierInvoiceNumber" Then
                                buyingSum = buyingSum + extraCharges2Buying
                                vatTotal = vatTotal + extraCharges2BuyVat
                            End If

                            'Extra Charges 3 Check + VAT If different inv num then split else if merge
                            If dtCharges.Tables("extraCharges").Rows(b).Item("generalSupplier3InvoiceNumber") <> "" And dtCharges.Tables("extraCharges").Rows(b).Item("generalSupplier3InvoiceNumber") = supInvNum And dsBuyingInvoices.Tables(0).Rows(y).Item("Names") = "generalSupplier3InvoiceNumber" Then
                                buyingSum = buyingSum + extraCharges3Buying
                                vatTotal = vatTotal + extraCharges3BuyVat
                                currency = dtCharges.Tables("extraCharges").Rows(0).Item("extraCharges3BuyingCurr")
                            ElseIf (dtCharges.Tables("extraCharges").Rows(0).Item("generalSupplier3InvoiceNumber") = "" Or dtCharges.Tables("extraCharges").Rows(b).Item("generalSupplier3InvoiceNumber") = supInvNum) And dsBuyingInvoices.Tables(0).Rows(y).Item("Names") = "supplierInvoiceNumber" Then
                                buyingSum = buyingSum + extraCharges3Buying
                                vatTotal = vatTotal + extraCharges3BuyVat
                            End If
                        End If
                    Next
                End If

                ' OLD VAT AND EXTRA CHARGES CALCULATIONS
                'vatTotal = vatTotal + buyVat + freightBuyVat + customsBuyVat + extraCharges1BuyVat + extraCharges2BuyVat + extraCharges3BuyVat
                'buyingSum = buyingSum + freightBuying + customsBuying + extraCharges1Buying + extraCharges2Buying + extraCharges3Buying

                'Total Price + VAT
                Dim grandTotalBuy As Double = 0.0
                grandTotalBuy = grandTotalBuy + buyingSum + vatTotal

                If dtData.Tables("outputvalues").Rows(0).Item("vesselName") = "NA" Then
                    dtData.Tables("outputvalues").Rows(0).Item("vesselName") = "NOT SPECIFIED-PO " + dtData.Tables("outputvalues").Rows(0).Item("custPurchOrderRef")
                End If

                Dim narrative = dtData.Tables("outputvalues").Rows(0).Item("narativeOfEntry") + "-" + dtData.Tables("outputvalues").Rows(0).Item("vesselName")

                Dim ifExist = chkIfAllExist(dtData.Tables("outputvalues").Rows(0).Item("supplierCode"))

                If vatTotal > 0 Then

                    ' Increment each time we insert a row in array
                    If u > 0 Then u += 1

                    'invoicesBuyArray(u, 0) = dsBuyingInvoices.Tables("invoicesSelling").Rows(y).Item("deliveryID")
                    invoicesBuyArray(u, 0) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("deliveryMergedID")
                    invoicesBuyArray(u, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                    If supInvNum.Contains("??INV") Or dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") <> "supplierInvoiceNumber" Then
                        If supInvNum.Contains("??INV") Then
                            invoicesBuyArray(u, 2) = "39002"
                            If dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "forwarderInvoiceNumber" Then
                                narrative = narrative + "- " + dtData.Tables("outputvalues").Rows(0).Item("forwarderCompName")
                            ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "clearingAgentInvoiceNumber" Then
                                narrative = narrative + "- " + dtData.Tables("outputvalues").Rows(0).Item("clearingAgentCompName")
                            ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "generalSupplier1InvoiceNumber" Then
                                narrative = narrative + "- " + dtData.Tables("outputvalues").Rows(0).Item("generalSupplier1CompName")
                            ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "generalSupplier2InvoiceNumber" Then
                                narrative = narrative + "- " + dtData.Tables("outputvalues").Rows(0).Item("generalSupplier2CompName")
                            ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "generalSupplier3InvoiceNumber" Then
                                narrative = narrative + "- " + dtData.Tables("outputvalues").Rows(0).Item("generalSupplier3CompName")
                            End If
                        Else
                            If dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "forwarderInvoiceNumber" Then
                                Try
                                    invoicesBuyArray(u, 2) = dsBuyingInvoices.Tables(0).Rows(y).Item("forwarderCode")
                                Catch ex As Exception

                                End Try
                                narrative = narrative + "- " + dtData.Tables("outputvalues").Rows(0).Item("forwarderCompName")
                            ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "clearingAgentInvoiceNumber" Then
                                Try
                                    invoicesBuyArray(u, 2) = dsBuyingInvoices.Tables(0).Rows(y).Item("clearingAgentCode")
                                Catch ex As Exception

                                End Try
                                narrative = narrative + "- " + dtData.Tables("outputvalues").Rows(0).Item("clearingAgentCompName")
                            ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "generalSupplier1InvoiceNumber" Then
                                Try
                                    invoicesBuyArray(u, 2) = dsBuyingInvoices.Tables(0).Rows(y).Item("generalSupplier1Code")
                                Catch ex As Exception

                                End Try
                                narrative = narrative + "- " + dtData.Tables("outputvalues").Rows(0).Item("generalSupplier1CompName")
                            ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "generalSupplier2InvoiceNumber" Then
                                Try
                                    invoicesBuyArray(u, 2) = dsBuyingInvoices.Tables(0).Rows(y).Item("generalSupplier2Code")
                                Catch ex As Exception

                                End Try
                                narrative = narrative + "- " + dtData.Tables("outputvalues").Rows(0).Item("generalSupplier2CompName")
                            ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "generalSupplier3InvoiceNumber" Then
                                Try
                                    invoicesBuyArray(u, 2) = dsBuyingInvoices.Tables(0).Rows(y).Item("generalSupplier3Code")
                                Catch ex As Exception

                                End Try
                                narrative = narrative + "- " + dtData.Tables("outputvalues").Rows(0).Item("generalSupplier3CompName")
                            End If
                        End If
                    Else
                        invoicesBuyArray(u, 2) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("supplierCode")
                    End If
                    If supInvNum.Contains("??INV") Then
                        invoicesBuyArray(u, 3) = "39002"
                    ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "forwarderInvoiceNumber" Then
                        If dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("deliveryMethod_ID") = 2 Then
                            invoicesBuyArray(u, 3) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("supplierCode")
                        Else
                            invoicesBuyArray(u, 3) = dsBuyingInvoices.Tables(0).Rows(y).Item("forwarderCode")
                        End If
                    ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "clearingAgentInvoiceNumber" Then
                        invoicesBuyArray(u, 3) = dsBuyingInvoices.Tables(0).Rows(y).Item("clearingAgentCode")
                    ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "generalSupplier1InvoiceNumber" Then
                        invoicesBuyArray(u, 3) = dsBuyingInvoices.Tables(0).Rows(y).Item("generalSupplier1Code")
                    ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "generalSupplier2InvoiceNumber" Then
                        invoicesBuyArray(u, 3) = dsBuyingInvoices.Tables(0).Rows(y).Item("generalSupplier2Code")
                    Else
                        invoicesBuyArray(u, 3) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("supplierCode")
                    End If
                    invoicesBuyArray(u, 4) = supInvNum
                    Try
                        If dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "forwarderInvoiceNumber" Then
                            invoicesBuyArray(u, 5) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("forwarderInvoiceDate")
                        ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "generalSupplier1InvoiceNumber" Then
                            invoicesBuyArray(u, 5) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("generalSupplier1InvoiceDate")
                        Else
                            invoicesBuyArray(u, 5) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("supplierInvoiceDate")
                        End If
                    Catch
                        MessageBox.Show("Supplier Invoice Number is missing from one of the deliveries you are trying to export.", "Important Message", MessageBoxButtons.OK)
                    End Try
                    invoicesBuyArray(u, 6) = TruncString.Truncate(narrative.ToString(), 60)
                    If dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "forwarderInvoiceNumber" Then
                        invoicesBuyArray(u, 7) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("dueDate2")
                    Else
                        Try
                            invoicesBuyArray(u, 7) = dtSupDueDate.Tables(0).Rows(0).Item("supplierDueDate2")
                        Catch ex As Exception
                            invoicesBuyArray(u, 7) = dtData.Tables("outputvalues").Rows(0).Item("supplierDueDate2")
                        End Try
                    End If
                    invoicesBuyArray(u, 8) = currency ' NEED SUPPLIER CURRENCY CODE
                    invoicesBuyArray(u, 9) = grandTotalBuy.ToString()
                    invoicesBuyArray(u, 10) = "0" 'VAT Number
                    invoicesBuyArray(u, 11) = "7" 'VAT Category
                    invoicesBuyArray(u, 12) = vatTotal.ToString()
                    If supInvNum.Contains("??INV") Then
                        invoicesBuyArray(u, 13) = "" 'INC_EXP_CATEGORY
                    Else
                        invoicesBuyArray(u, 13) = "2.1" 'INC_EXP_CATEGORY
                    End If
                    invoicesBuyArray(u, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                    invoicesBuyArray(u, 15) = dtData.Tables("outputvalues").Rows(0).Item("invoiceDate")
                    invoicesBuyArray(u, 16) = danaosUsername
                    invoicesBuyArray(u, 17) = "C" 'DEBIT CREDI
                    invoicesBuyArray(u, 18) = "" 'VESSEL_CODE
                    Try
                        invoicesBuyArray(u, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                    Catch ex As Exception
                        invoicesBuyArray(u, 19) = dtData.Tables("outputvalues").Rows(0).Item("refNum")
                    End Try
                    invoicesBuyArray(u, 20) = "023" 'JOURNAL_COMPANY
                    invoicesBuyArray(u, 21) = "40" 'JOURNALSERIES
                    invoicesBuyArray(u, 22) = "0" ' BOOK VALUE
                    invoicesBuyArray(u, 23) = "0" ' ENTRYRATE
                    If ifExist = True Then
                        invoicesBuyArray(u, 24) = "True"
                    Else
                        invoicesBuyArray(u, 24) = "False"
                    End If

                    ' insert total + vat
                    u += 1
                    'Inserts the First Row of the Amount without VAT
                    'invoicesBuyArray(u, 0) = dsBuyingInvoices.Tables("invoicesSelling").Rows(y).Item("deliveryID")
                    invoicesBuyArray(u, 0) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("deliveryMergedID")
                    invoicesBuyArray(u, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                    invoicesBuyArray(u, 2) = dtData.Tables("outputvalues").Rows(0).Item("purchaseCode")
                    If supInvNum.Contains("??INV") Then
                        invoicesBuyArray(u, 3) = "39002"
                    ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "forwarderInvoiceNumber" Then
                        If dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("deliveryMethod_ID") = 2 Then
                            invoicesBuyArray(u, 3) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("supplierCode")
                        Else
                            invoicesBuyArray(u, 3) = dsBuyingInvoices.Tables(0).Rows(y).Item("forwarderCode")
                        End If
                    ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "clearingAgentInvoiceNumber" Then
                        invoicesBuyArray(u, 3) = dsBuyingInvoices.Tables(0).Rows(y).Item("clearingAgentCode")
                    ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "generalSupplier1InvoiceNumber" Then
                        invoicesBuyArray(u, 3) = dsBuyingInvoices.Tables(0).Rows(y).Item("generalSupplier1Code")
                    ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "generalSupplier2InvoiceNumber" Then
                        invoicesBuyArray(u, 3) = dsBuyingInvoices.Tables(0).Rows(y).Item("generalSupplier2Code")
                    Else
                        invoicesBuyArray(u, 3) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("supplierCode")
                    End If
                    invoicesBuyArray(u, 4) = supInvNum
                    Try
                        If dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "forwarderInvoiceNumber" Then
                            invoicesBuyArray(u, 5) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("forwarderInvoiceDate")
                        ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "generalSupplier1InvoiceNumber" Then
                            invoicesBuyArray(u, 5) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("generalSupplier1InvoiceDate")
                        Else
                            invoicesBuyArray(u, 5) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("supplierInvoiceDate")
                        End If
                    Catch
                        MessageBox.Show("Supplier Invoice Number is missing from one of the deliveries you are trying to export.", "Important Message", MessageBoxButtons.OK)
                    End Try
                    invoicesBuyArray(u, 6) = TruncString.Truncate(narrative.ToString(), 60)
                    If dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "forwarderInvoiceNumber" Then
                        invoicesBuyArray(u, 7) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("dueDate2")
                    Else
                        Try
                            invoicesBuyArray(u, 7) = dtSupDueDate.Tables(0).Rows(0).Item("supplierDueDate2")
                        Catch ex As Exception
                            invoicesBuyArray(u, 7) = dtData.Tables("outputvalues").Rows(0).Item("supplierDueDate2")
                        End Try
                    End If
                    invoicesBuyArray(u, 8) = currency ' NEED SUPPLIER CURRENCY CODE
                    invoicesBuyArray(u, 9) = buyingSum.ToString()
                    invoicesBuyArray(u, 10) = "0" 'VAT Number
                    invoicesBuyArray(u, 11) = "2" 'VAT Category
                    invoicesBuyArray(u, 12) = "0"
                    invoicesBuyArray(u, 13) = "" 'INC_EXP_CATEGORY
                    invoicesBuyArray(u, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                    invoicesBuyArray(u, 15) = dtData.Tables("outputvalues").Rows(0).Item("invoiceDate")
                    invoicesBuyArray(u, 16) = danaosUsername
                    invoicesBuyArray(u, 17) = "D" 'DEBIT CREDI
                    invoicesBuyArray(u, 18) = "" 'VESSEL_CODE
                    Try
                        invoicesBuyArray(u, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                    Catch ex As Exception
                        invoicesBuyArray(u, 19) = dtData.Tables("outputvalues").Rows(0).Item("refNum")
                    End Try
                    invoicesBuyArray(u, 20) = "023" 'JOURNAL_COMPANY
                    invoicesBuyArray(u, 21) = "40" 'JOURNALSERIES
                    invoicesBuyArray(u, 22) = "0" ' BOOK VALUE
                    invoicesBuyArray(u, 23) = "0" ' ENTRYRATE
                    If ifExist = True Then
                        invoicesBuyArray(u, 24) = "True"
                    Else
                        invoicesBuyArray(u, 24) = "False"
                    End If

                    ' if vat is 0 skip dont insert anything
                    u += 1
                    'invoicesBuyArray(u, 0) = dsBuyingInvoices.Tables("invoicesSelling").Rows(y).Item("deliveryID")
                    invoicesBuyArray(u, 0) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("deliveryMergedID")
                    invoicesBuyArray(u, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                    invoicesBuyArray(u, 2) = "38601"
                    invoicesBuyArray(u, 3) = dtData.Tables("outputvalues").Rows(0).Item("purchaseCode")
                    invoicesBuyArray(u, 4) = supInvNum
                    Try
                        If dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "forwarderInvoiceNumber" Then
                            invoicesBuyArray(u, 5) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("forwarderInvoiceDate")
                        ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "generalSupplier1InvoiceNumber" Then
                            invoicesBuyArray(u, 5) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("generalSupplier1InvoiceDate")
                        Else
                            invoicesBuyArray(u, 5) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("supplierInvoiceDate")
                        End If
                    Catch
                        MessageBox.Show("Supplier Invoice Number is missing from one of the deliveries you are trying to export.", "Important Message", MessageBoxButtons.OK)
                    End Try
                    invoicesBuyArray(u, 6) = TruncString.Truncate(narrative.ToString(), 60)
                    If dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "forwarderInvoiceNumber" Then
                        invoicesBuyArray(u, 7) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("dueDate2")
                    Else
                        Try
                            invoicesBuyArray(u, 7) = dtSupDueDate.Tables(0).Rows(0).Item("supplierDueDate2")
                        Catch ex As Exception
                            invoicesBuyArray(u, 7) = dtData.Tables("outputvalues").Rows(0).Item("supplierDueDate2")
                        End Try
                    End If
                    invoicesBuyArray(u, 8) = currency ' NEED SUPPLIER CURRENCY CODE
                    invoicesBuyArray(u, 9) = vatTotal.ToString()
                    invoicesBuyArray(u, 10) = "0" 'VAT Number
                    invoicesBuyArray(u, 11) = "2" 'VAT Category
                    invoicesBuyArray(u, 12) = "0"
                    invoicesBuyArray(u, 13) = "" 'INC_EXP_CATEGORY
                    invoicesBuyArray(u, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                    invoicesBuyArray(u, 15) = dtData.Tables("outputvalues").Rows(0).Item("invoiceDate")
                    invoicesBuyArray(u, 16) = danaosUsername
                    invoicesBuyArray(u, 17) = "D" 'DEBIT CREDI
                    invoicesBuyArray(u, 18) = "" 'VESSEL_CODE
                    Try
                        invoicesBuyArray(u, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                    Catch ex As Exception
                        invoicesBuyArray(u, 19) = dtData.Tables("outputvalues").Rows(0).Item("refNum")
                    End Try
                    invoicesBuyArray(u, 20) = "023" 'JOURNAL_COMPANY
                    invoicesBuyArray(u, 21) = "40" 'JOURNALSERIES
                    invoicesBuyArray(u, 22) = "0" ' BOOK VALUE                        
                    invoicesBuyArray(u, 23) = "0" ' ENTRYRATE
                    If ifExist = True Then
                        invoicesBuyArray(u, 24) = "True"
                    Else
                        invoicesBuyArray(u, 24) = "False"
                    End If
                Else
                    ' Increment each time we insert a row in array
                    If u > 0 Then u += 1

                    'invoicesBuyArray(u, 0) = dsBuyingInvoices.Tables("invoicesSelling").Rows(y).Item("deliveryID")
                    invoicesBuyArray(u, 0) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("deliveryMergedID")
                    invoicesBuyArray(u, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                    invoicesBuyArray(u, 2) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("supplierCode")
                    If supInvNum.Contains("??INV") Or dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") <> "supplierInvoiceNumber" Then
                        If supInvNum.Contains("??INV") Then
                            invoicesBuyArray(u, 2) = "39002"
                            If dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "forwarderInvoiceNumber" Then
                                narrative = narrative + "- " + dtData.Tables("outputvalues").Rows(0).Item("forwarderCompName")
                            ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "clearingAgentInvoiceNumber" Then
                                narrative = narrative + "- " + dtData.Tables("outputvalues").Rows(0).Item("clearingAgentCompName")
                            ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "generalSupplier1InvoiceNumber" Then
                                narrative = narrative + "- " + dtData.Tables("outputvalues").Rows(0).Item("generalSupplier1CompName")
                            ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "generalSupplier2InvoiceNumber" Then
                                narrative = narrative + "- " + dtData.Tables("outputvalues").Rows(0).Item("generalSupplier2CompName")
                            ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "generalSupplier3InvoiceNumber" Then
                                narrative = narrative + "- " + dtData.Tables("outputvalues").Rows(0).Item("generalSupplier3CompName")
                            End If
                        Else
                            If dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "forwarderInvoiceNumber" Then
                                Try
                                    invoicesBuyArray(u, 2) = dsBuyingInvoices.Tables(0).Rows(y).Item("forwarderCode")
                                Catch ex As Exception

                                End Try
                                narrative = narrative + "- " + dtData.Tables("outputvalues").Rows(0).Item("forwarderCompName")
                            ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "clearingAgentInvoiceNumber" Then
                                Try
                                    invoicesBuyArray(u, 2) = dsBuyingInvoices.Tables(0).Rows(y).Item("clearingAgentCode")
                                Catch ex As Exception

                                End Try
                                narrative = narrative + "- " + dtData.Tables("outputvalues").Rows(0).Item("clearingAgentCompName")
                            ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "generalSupplier1InvoiceNumber" Then
                                Try
                                    invoicesBuyArray(u, 2) = dsBuyingInvoices.Tables(0).Rows(y).Item("generalSupplier1Code")
                                Catch ex As Exception

                                End Try
                                narrative = narrative + "- " + dtData.Tables("outputvalues").Rows(0).Item("generalSupplier1CompName")
                            ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "generalSupplier2InvoiceNumber" Then
                                Try
                                    invoicesBuyArray(u, 2) = dsBuyingInvoices.Tables(0).Rows(y).Item("generalSupplier2Code")
                                Catch ex As Exception

                                End Try
                                narrative = narrative + "- " + dtData.Tables("outputvalues").Rows(0).Item("generalSupplier2CompName")
                            ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "generalSupplier3InvoiceNumber" Then
                                Try
                                    invoicesBuyArray(u, 2) = dsBuyingInvoices.Tables(0).Rows(y).Item("generalSupplier3Code")
                                Catch ex As Exception

                                End Try
                                narrative = narrative + "- " + dtData.Tables("outputvalues").Rows(0).Item("generalSupplier3CompName")
                            End If
                        End If
                    Else
                        invoicesBuyArray(u, 2) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("supplierCode")
                    End If
                    If supInvNum.Contains("??INV") Then
                        invoicesBuyArray(u, 3) = "39002"
                    ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "forwarderInvoiceNumber" Then
                        If dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("deliveryMethod_ID") = 2 Then
                            invoicesBuyArray(u, 3) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("supplierCode")
                        Else
                            invoicesBuyArray(u, 3) = dsBuyingInvoices.Tables(0).Rows(y).Item("forwarderCode")
                        End If
                    ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "clearingAgentInvoiceNumber" Then
                        invoicesBuyArray(u, 3) = dsBuyingInvoices.Tables(0).Rows(y).Item("clearingAgentCode")
                    ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "generalSupplier1InvoiceNumber" Then
                        invoicesBuyArray(u, 3) = dsBuyingInvoices.Tables(0).Rows(y).Item("generalSupplier1Code")
                    ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "generalSupplier2InvoiceNumber" Then
                        invoicesBuyArray(u, 3) = dsBuyingInvoices.Tables(0).Rows(y).Item("generalSupplier2Code")
                    Else
                        invoicesBuyArray(u, 3) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("supplierCode")
                    End If
                    invoicesBuyArray(u, 4) = supInvNum
                    Try
                        If dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "forwarderInvoiceNumber" Then
                            invoicesBuyArray(u, 5) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("forwarderInvoiceDate")
                        ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "generalSupplier1InvoiceNumber" Then
                            invoicesBuyArray(u, 5) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("generalSupplier1InvoiceDate")
                        Else
                            invoicesBuyArray(u, 5) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("supplierInvoiceDate")
                        End If
                    Catch
                        MessageBox.Show("Supplier Invoice Number is missing from one of the deliveries you are trying to export.", "Important Message", MessageBoxButtons.OK)
                    End Try
                    invoicesBuyArray(u, 6) = TruncString.Truncate(narrative.ToString(), 60)
                    If dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "forwarderInvoiceNumber" Then
                        invoicesBuyArray(u, 7) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("dueDate2")
                    Else
                        Try
                            invoicesBuyArray(u, 7) = dtSupDueDate.Tables(0).Rows(0).Item("supplierDueDate2")
                        Catch ex As Exception
                            invoicesBuyArray(u, 7) = dtData.Tables("outputvalues").Rows(0).Item("supplierDueDate2")
                        End Try
                    End If
                    invoicesBuyArray(u, 8) = currency ' NEED SUPPLIER CURRENCY CODE
                    invoicesBuyArray(u, 9) = grandTotalBuy.ToString()
                    invoicesBuyArray(u, 10) = "0" 'VAT Number
                    invoicesBuyArray(u, 11) = "2" 'VAT Category
                    invoicesBuyArray(u, 12) = vatTotal.ToString()
                    If supInvNum.Contains("??INV") Then
                        invoicesBuyArray(u, 13) = "" 'INC_EXP_CATEGORY
                    Else
                        invoicesBuyArray(u, 13) = "2.1" 'INC_EXP_CATEGORY
                    End If
                    invoicesBuyArray(u, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                    invoicesBuyArray(u, 15) = dtData.Tables("outputvalues").Rows(0).Item("invoiceDate")
                    invoicesBuyArray(u, 16) = danaosUsername
                    invoicesBuyArray(u, 17) = "C" 'DEBIT CREDI
                    invoicesBuyArray(u, 18) = "" 'VESSEL_CODE
                    Try
                        invoicesBuyArray(u, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                    Catch ex As Exception
                        invoicesBuyArray(u, 19) = dtData.Tables("outputvalues").Rows(0).Item("refNum")
                    End Try
                    invoicesBuyArray(u, 20) = "023" 'JOURNAL_COMPANY
                    invoicesBuyArray(u, 21) = "40" 'JOURNALSERIES
                    invoicesBuyArray(u, 22) = "0" ' BOOK VALUE
                    invoicesBuyArray(u, 23) = "0" ' ENTRYRATE
                    If ifExist = True Then
                        invoicesBuyArray(u, 24) = "True"
                    Else
                        invoicesBuyArray(u, 24) = "False"
                    End If

                    ' insert total + vat
                    u += 1
                    'Inserts the First Row of the Amount without VAT
                    'invoicesBuyArray(u, 0) = dsBuyingInvoices.Tables("invoicesSelling").Rows(y).Item("deliveryID")
                    invoicesBuyArray(u, 0) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("deliveryMergedID")
                    invoicesBuyArray(u, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                    invoicesBuyArray(u, 2) = dtData.Tables("outputvalues").Rows(0).Item("purchaseCode")
                    If supInvNum.Contains("??INV") Then
                        invoicesBuyArray(u, 3) = "39002"
                    ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "forwarderInvoiceNumber" Then
                        If dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("deliveryMethod_ID") = 2 Then
                            invoicesBuyArray(u, 3) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("supplierCode")
                        Else
                            invoicesBuyArray(u, 3) = dsBuyingInvoices.Tables(0).Rows(y).Item("forwarderCode")
                        End If
                    ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "clearingAgentInvoiceNumber" Then
                        invoicesBuyArray(u, 3) = dsBuyingInvoices.Tables(0).Rows(y).Item("clearingAgentCode")
                    ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "generalSupplier1InvoiceNumber" Then
                        invoicesBuyArray(u, 3) = dsBuyingInvoices.Tables(0).Rows(y).Item("generalSupplier1Code")
                    ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "generalSupplier2InvoiceNumber" Then
                        invoicesBuyArray(u, 3) = dsBuyingInvoices.Tables(0).Rows(y).Item("generalSupplier2Code")
                    Else
                        invoicesBuyArray(u, 3) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("supplierCode")
                    End If
                    invoicesBuyArray(u, 4) = supInvNum
                    Try
                        If dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "forwarderInvoiceNumber" Then
                            invoicesBuyArray(u, 5) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("forwarderInvoiceDate")
                        ElseIf dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "generalSupplier1InvoiceNumber" Then
                            invoicesBuyArray(u, 5) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("generalSupplier1InvoiceDate")
                        Else
                            invoicesBuyArray(u, 5) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("supplierInvoiceDate")
                        End If
                    Catch
                        MessageBox.Show("Supplier Invoice Number is missing from one of the deliveries you are trying to export.", "Important Message", MessageBoxButtons.OK)
                    End Try
                    invoicesBuyArray(u, 6) = TruncString.Truncate(narrative.ToString(), 60)
                    If dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("Names") = "forwarderInvoiceNumber" Then
                        invoicesBuyArray(u, 7) = dsBuyingInvoices.Tables("invoicesBuying").Rows(y).Item("dueDate2")
                    Else
                        Try
                            invoicesBuyArray(u, 7) = dtSupDueDate.Tables(0).Rows(0).Item("supplierDueDate2")
                        Catch ex As Exception
                            invoicesBuyArray(u, 7) = dtData.Tables("outputvalues").Rows(0).Item("supplierDueDate2")
                        End Try
                    End If
                    invoicesBuyArray(u, 8) = currency ' NEED SUPPLIER CURRENCY CODE
                    invoicesBuyArray(u, 9) = buyingSum.ToString()
                    invoicesBuyArray(u, 10) = "0" 'VAT Number
                    invoicesBuyArray(u, 11) = "2" 'VAT Category
                    invoicesBuyArray(u, 12) = vatTotal.ToString()
                    invoicesBuyArray(u, 13) = "" 'INC_EXP_CATEGORY
                    invoicesBuyArray(u, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                    invoicesBuyArray(u, 15) = dtData.Tables("outputvalues").Rows(0).Item("invoiceDate")
                    invoicesBuyArray(u, 16) = danaosUsername
                    invoicesBuyArray(u, 17) = "D" 'DEBIT CREDI
                    invoicesBuyArray(u, 18) = "" 'VESSEL_CODE
                    Try
                        invoicesBuyArray(u, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                    Catch ex As Exception
                        invoicesBuyArray(u, 19) = dtData.Tables("outputvalues").Rows(0).Item("refNum")
                    End Try
                    invoicesBuyArray(u, 20) = "023" 'JOURNAL_COMPANY
                    invoicesBuyArray(u, 21) = "40" 'JOURNALSERIES
                    invoicesBuyArray(u, 22) = "0" ' BOOK VALUE                        
                    invoicesBuyArray(u, 23) = "0" ' ENTRYRATE
                    If ifExist = True Then
                        invoicesBuyArray(u, 24) = "True"
                    Else
                        invoicesBuyArray(u, 24) = "False"
                    End If

                End If
            End If

        Next

        Return invoicesBuyArray
        'END BUYING
        '====================================================================================================================='
    End Function

    ''' <summary>
    ''' Credit Note 
    ''' </summary>
    ''' <returns></returns>
    Public Function creditNote() As String(,)
        '======================================================================================================================'
        'CREDIT NOTE
        'Send the OrderID to get the Credit Note
        Dim dsCreditNote = generaldataset.creditNoteType(orderID)
        Dim rowsCount = dsCreditNote.Tables(0).Rows.Count - 1
        Dim invRowsCount = (rowsCount + 1) * 3
        Dim creditNoteArray(invRowsCount, 24) As String
        Dim dsCustomerCN As DataSet
        Dim dsBrokerCN As DataSet
        Dim dsSupplierCN As DataSet
        Dim y As Integer = 0
        Dim mergedID = 0
        Dim invNum = ""

        ' CHeck if there is credit in order
        If dsCreditNote.Tables(0).Rows.Count > 0 Then

            ' Loop to get creditNote
            For i = 0 To dsCreditNote.Tables(0).Rows.Count - 1

                If dsCreditNote.Tables(0).Rows(i).Item("exported") = 0 Then

                    ' Check creditNote Type
                    If dsCreditNote.Tables(0).Rows(i).Item("creditNoteType_ID") = 1 Then

                        'get the merged value find the customer CN items
                        mergedID = dsCreditNote.Tables(0).Rows(i).Item("deliveryMergedID")
                        dsCustomerCN = generaldataset.creditNoteCustomer(orderID, mergedID)
                        Dim totalCustomerCN = 0.0
                        Dim multi = 0.0

                        ' Loop to calculate Customer CN items
                        For o = 0 To dsCustomerCN.Tables(0).Rows.Count - 1
                            If Not IsDBNull(dsCustomerCN.Tables(0).Rows(o).Item("creditNoteDelQty")) And Not IsDBNull(dsCustomerCN.Tables(0).Rows(o).Item("cnCustItemPrice")) Then
                                multi = dsCustomerCN.Tables(0).Rows(o).Item("creditNoteDelQty") * dsCustomerCN.Tables(0).Rows(o).Item("cnCustItemPrice")
                            End If
                            If Not IsDBNull(dsCustomerCN.Tables(0).Rows(o).Item("creditNoteCustLSAmount")) Then
                                totalCustomerCN = totalCustomerCN + dsCustomerCN.Tables(0).Rows(o).Item("creditNoteCustLSAmount")
                            End If
                            If totalCustomerCN = 0 Then
                                totalCustomerCN = totalCustomerCN + Math.Round(multi, 2, MidpointRounding.AwayFromZero)
                            End If
                        Next

                        ' APply customer CN discount
                        For o = 0 To dsCustomerCN.Tables(0).Rows.Count - 1
                            If dsCustomerCN.Tables(0).Rows(o).Item("creditNoteDiscountLS") <> 0 Then
                                totalCustomerCN = totalCustomerCN - dsCustomerCN.Tables(0).Rows(o).Item("creditNoteDiscountLS")
                                GoTo exitDiscount
                            End If
                        Next

exitDiscount:

                        ' Get the output values
                        Dim dtData = generaldataset.loadOutputValues(orderID, mergedID)

                        ' Set NA to Not Specified
                        If dtData.Tables("outputvalues").Rows(0).Item("vesselName") = "NA" Then
                            dtData.Tables("outputvalues").Rows(0).Item("vesselName") = "NOT SPECIFIED"
                        End If

                        ' Narrative
                        Dim narrative = dtData.Tables("outputvalues").Rows(0).Item("narativeOfEntry") + "-" + dtData.Tables("outputvalues").Rows(0).Item("vesselName") + "-CREDIT NOTE"

                        If y > 0 Then y += 1

                        'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                        creditNoteArray(y, 0) = dsCreditNote.Tables(0).Rows(i).Item("deliveryMergedID")
                        creditNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                        creditNoteArray(y, 2) = dtData.Tables("outputvalues").Rows(0).Item("customerCode")
                        creditNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("customerCode")
                        creditNoteArray(y, 4) = dsCreditNote.Tables(0).Rows(i).Item("creditNoteNumber")
                        creditNoteArray(y, 5) = dsCreditNote.Tables(0).Rows(i).Item("creditNoteDate")
                        creditNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                        creditNoteArray(y, 7) = dsCreditNote.Tables(0).Rows(i).Item("dueDate2")
                        creditNoteArray(y, 8) = dtData.Tables("outputvalues").Rows(0).Item("customerCurrencyCode")
                        creditNoteArray(y, 9) = totalCustomerCN.ToString()
                        creditNoteArray(y, 10) = "0" 'VAT Number
                        creditNoteArray(y, 11) = "2" 'VAT Category
                        creditNoteArray(y, 12) = "0"
                        creditNoteArray(y, 13) = "1.1" 'INC_EXP_CATEGORY
                        creditNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                        creditNoteArray(y, 15) = dsCreditNote.Tables(0).Rows(i).Item("creditNoteDate")
                        creditNoteArray(y, 16) = danaosUsername
                        creditNoteArray(y, 17) = "C" 'DEBIT CREDI
                        creditNoteArray(y, 18) = "" 'VESSEL_CODE
                        Try
                            creditNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                        Catch ex As Exception
                            creditNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("refNum")
                        End Try
                        creditNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                        creditNoteArray(y, 21) = "40" 'JOURNALSERIES
                        creditNoteArray(y, 22) = "0" ' BOOK VALUE
                        creditNoteArray(y, 23) = "0" ' ENTRYRATE
                        creditNoteArray(y, 24) = "True" ' ENTRYRATE

                        ' Adds a Next Row of the Amount + VAT
                        y += 1
                        'Inserts the First Row of the Amount without VAT
                        'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                        creditNoteArray(y, 0) = dsCreditNote.Tables(0).Rows(i).Item("deliveryMergedID")
                        creditNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                        creditNoteArray(y, 2) = dtData.Tables("outputvalues").Rows(0).Item("salesCode")
                        creditNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("customerCode")
                        creditNoteArray(y, 4) = dsCreditNote.Tables(0).Rows(i).Item("creditNoteNumber")
                        creditNoteArray(y, 5) = dsCreditNote.Tables(0).Rows(i).Item("creditNoteDate")
                        creditNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                        creditNoteArray(y, 7) = dsCreditNote.Tables(0).Rows(i).Item("dueDate2")
                        creditNoteArray(y, 8) = dtData.Tables("outputvalues").Rows(0).Item("customerCurrencyCode")
                        creditNoteArray(y, 9) = totalCustomerCN.ToString()
                        creditNoteArray(y, 10) = "0" 'VAT Number
                        creditNoteArray(y, 11) = "2" 'VAT Category
                        creditNoteArray(y, 12) = "0"
                        creditNoteArray(y, 13) = "" 'INC_EXP_CATEGORY
                        creditNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                        creditNoteArray(y, 15) = dsCreditNote.Tables(0).Rows(i).Item("creditNoteDate")
                        creditNoteArray(y, 16) = danaosUsername
                        creditNoteArray(y, 17) = "D" 'DEBIT CREDI
                        creditNoteArray(y, 18) = "" 'VESSEL_CODE
                        Try
                            creditNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                        Catch ex As Exception
                            creditNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("refNum")
                        End Try
                        creditNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                        creditNoteArray(y, 21) = "40" 'JOURNALSERIES
                        creditNoteArray(y, 22) = "0" ' BOOK VALUE
                        creditNoteArray(y, 23) = "0" ' ENTRYRATE
                        creditNoteArray(y, 24) = "True" ' ENTRYRATE




                        ' If CN is Brokerage (Type 2)
                    ElseIf dsCreditNote.Tables(0).Rows(i).Item("creditNoteType_ID") = 2 Then

                        Dim itemsCalc = 0.0
                        Dim totalBrokerCN = 0.0
                        Dim deliveryIDParam = dsCreditNote.Tables(0).Rows(i).Item("deliveryID")

                        mergedID = dsCreditNote.Tables(0).Rows(i).Item("deliveryMergedID")
                        Dim creditNoteTMP = dsCreditNote.Tables(0).Rows(i).Item("creditNoteNumber")
                        Dim invoiceForCNBrokerage = generaldataset.invoiceForCNBroker(deliveryIDParam)
                        Dim delBroker = generaldataset.deliveryBrokerCode(creditNoteTMP)
                        invNum = invoiceForCNBrokerage.Tables(0).Rows(0).Item("invoiceNumber")

                        ' Call selling items and broker CN LS or Per
                        Dim dsSellingCalc = generaldataset.loadItemsSelling(orderID, mergedID, invNum)
                        dsBrokerCN = generaldataset.creditNoteBroker(orderID, mergedID, delBroker.Tables(0).Rows(0).Item("deliveryBroker_ID"))
                        Dim dtExchange = generaldataset.loadExchangeRates(orderID)

                        If dsBrokerCN.Tables(0).Rows(0).Item("brokerageCurrencyCode") = dtExchange.Tables(0).Rows(0).Item("supplierCurrencyCode") Then
                            ' Loop to get the items price
                            For o = 0 To dsSellingCalc.Tables(0).Rows.Count - 1
                                Dim multiQtyPr = dsSellingCalc.Tables("itemsSelling").Rows(o).Item("quotItemPrice") + (dsSellingCalc.Tables("itemsSelling").Rows(o).Item("quotItemPrice") * dsSellingCalc.Tables("itemsSelling").Rows(o).Item("quotItemIncrement") / 100)
                                For b = 0 To dtExchange.Tables(0).Rows.Count - 1
                                    If dsSellingCalc.Tables("itemsSelling").Rows(o).Item("supplierInvoiceNumber") = dtExchange.Tables(0).Rows(b).Item("supplierInvoiceNumber") Then
                                        multiQtyPr = Math.Round(multiQtyPr / dtExchange.Tables(0).Rows(b).Item("supplierExchangeRate") * dsBrokerCN.Tables(0).Rows(0).Item("brokerageExchangeRate"), 2)
                                    End If
                                Next
                                Dim idk = dsSellingCalc.Tables("itemsSelling").Rows(o).Item("deliveryDeliveredQty") * Math.Round(multiQtyPr, 2, MidpointRounding.AwayFromZero)
                                itemsCalc = itemsCalc + Math.Round(idk, 2)
                            Next
                        Else
                            Dim rate = ""

                            For o = 0 To dsSellingCalc.Tables(0).Rows.Count - 1
                                For b = 0 To dtExchange.Tables(0).Rows.Count - 1
                                    If dsSellingCalc.Tables("itemsSelling").Rows(o).Item("supplierInvoiceNumber") = dtExchange.Tables(0).Rows(b).Item("supplierInvoiceNumber") Then
                                        rate = dtExchange.Tables(0).Rows(b).Item("supplierExchangeRate")

                                        Dim exchanged = dsSellingCalc.Tables(0).Rows(o).Item("quotItemPrice") / rate * dsBrokerCN.Tables(0).Rows(0).Item("brokerageExchangeRate")
                                        Dim multiQtyPr = exchanged + Math.Round(exchanged * dsSellingCalc.Tables("itemsSelling").Rows(o).Item("quotItemIncrement") / 100, 2)
                                        Dim idk = dsSellingCalc.Tables("itemsSelling").Rows(o).Item("deliveryDeliveredQty") * multiQtyPr
                                        itemsCalc = itemsCalc + Math.Round(idk, 2)
                                    End If
                                Next
                            Next
                        End If

                        Dim what = 0.0

                        ' Add the Brokerage Fee to the items total
                        Try
                            If i = 1 Then
                                If dsBrokerCN.Tables(0).Rows(1).Item("deliveryBrokerageFeeLS") <> 0 Then
                                    totalBrokerCN = totalBrokerCN + dsBrokerCN.Tables(0).Rows(1).Item("deliveryBrokerageFeeLS")
                                ElseIf dsBrokerCN.Tables(0).Rows(1).Item("deliveryBrokerageFeePer") <> 0 Then
                                    what = itemsCalc - (itemsCalc * dsSellingCalc.Tables("itemsSelling").Rows(0).Item("customerDiscountPer") / 100)
                                    'what = itemsCalc * dsBrokerCN.Tables(0).Rows(0).Item("deliveryBrokerageFeePer") / 100
                                    totalBrokerCN = totalBrokerCN + (what * dsBrokerCN.Tables(0).Rows(1).Item("deliveryBrokerageFeePer") / 100)
                                    'totalBrokerCN = totalBrokerCN - Math.Round(totalBrokerCN * dsSellingCalc.Tables("itemsSelling").Rows(0).Item("customerDiscountPer") / 100, 2)
                                End If
                            ElseIf i = 0 Then
                                If dsBrokerCN.Tables(0).Rows(0).Item("deliveryBrokerageFeeLS") <> 0 Then
                                    totalBrokerCN = totalBrokerCN + dsBrokerCN.Tables(0).Rows(0).Item("deliveryBrokerageFeeLS")
                                ElseIf dsBrokerCN.Tables(0).Rows(0).Item("deliveryBrokerageFeePer") <> 0 Then
                                    what = itemsCalc - dsSellingCalc.Tables("itemsSelling").Rows(0).Item("customerDiscountLS") - (itemsCalc * dsSellingCalc.Tables("itemsSelling").Rows(0).Item("customerDiscountPer") / 100)
                                    'what = itemsCalc * dsBrokerCN.Tables(0).Rows(0).Item("deliveryBrokerageFeePer") / 100
                                    totalBrokerCN = Math.Round(totalBrokerCN + (what * dsBrokerCN.Tables(0).Rows(0).Item("deliveryBrokerageFeePer") / 100), 2)
                                    'totalBrokerCN = totalBrokerCN - Math.Round(totalBrokerCN * dsSellingCalc.Tables("itemsSelling").Rows(0).Item("customerDiscountPer") / 100, 2)
                                End If
                            ElseIf i = 2 Then
                                If dsBrokerCN.Tables(0).Rows(2).Item("deliveryBrokerageFeeLS") <> 0 Then
                                    totalBrokerCN = totalBrokerCN + dsBrokerCN.Tables(0).Rows(2).Item("deliveryBrokerageFeeLS")
                                ElseIf dsBrokerCN.Tables(0).Rows(2).Item("deliveryBrokerageFeePer") <> 0 Then
                                    what = itemsCalc - (itemsCalc * dsSellingCalc.Tables("itemsSelling").Rows(0).Item("customerDiscountPer") / 100)
                                    'what = itemsCalc * dsBrokerCN.Tables(0).Rows(0).Item("deliveryBrokerageFeePer") / 100
                                    totalBrokerCN = totalBrokerCN + (what * dsBrokerCN.Tables(0).Rows(2).Item("deliveryBrokerageFeePer") / 100)
                                    'totalBrokerCN = totalBrokerCN - Math.Round(totalBrokerCN * dsSellingCalc.Tables("itemsSelling").Rows(0).Item("customerDiscountPer") / 100, 2)
                                End If
                            End If
                        Catch ex As Exception
                            If dsBrokerCN.Tables(0).Rows(0).Item("deliveryBrokerageFeeLS") <> 0 Then
                                totalBrokerCN = totalBrokerCN + dsBrokerCN.Tables(0).Rows(0).Item("deliveryBrokerageFeeLS")
                            ElseIf dsBrokerCN.Tables(0).Rows(0).Item("deliveryBrokerageFeePer") <> 0 Then
                                what = itemsCalc - (itemsCalc * dsSellingCalc.Tables("itemsSelling").Rows(0).Item("customerDiscountPer") / 100)
                                'what = itemsCalc * dsBrokerCN.Tables(0).Rows(0).Item("deliveryBrokerageFeePer") / 100
                                totalBrokerCN = totalBrokerCN + (what * dsBrokerCN.Tables(0).Rows(0).Item("deliveryBrokerageFeePer") / 100)
                                'totalBrokerCN = totalBrokerCN - Math.Round(totalBrokerCN * dsSellingCalc.Tables("itemsSelling").Rows(0).Item("customerDiscountPer") / 100, 2)
                            End If

                        End Try

                        Dim dtData = generaldataset.loadOutputValues(orderID, mergedID)

                        If dtData.Tables("outputvalues").Rows(0).Item("vesselName") = "NA" Then
                            dtData.Tables("outputvalues").Rows(0).Item("vesselName") = "NOT SPECIFIED"
                        End If

                        Dim narrative = dtData.Tables("outputvalues").Rows(0).Item("narativeOfEntry") + "-" + dtData.Tables("outputvalues").Rows(0).Item("vesselName") + "-CREDIT NOTE"

                        Dim brokerCodeExists = False

                        brokerCodeExists = chkIfAllExist(dsBrokerCN.Tables(0).Rows(0).Item("brokerCode"))

                        If brokerCodeExists = True Then

                            If y > 0 Then y += 1

                            'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                            creditNoteArray(y, 0) = dsCreditNote.Tables(0).Rows(i).Item("deliveryMergedID")
                            creditNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                            'If i = 1 Then
                            '    creditNoteArray(y, 2) = dsBrokerCN.Tables(0).Rows(1).Item("brokerCode") 'BROKER 
                            '    creditNoteArray(y, 3) = dsBrokerCN.Tables(0).Rows(1).Item("brokerCode") ' BROKER
                            'ElseIf i = 0 Then
                            creditNoteArray(y, 2) = dsBrokerCN.Tables(0).Rows(0).Item("brokerCode") 'BROKER 
                            creditNoteArray(y, 3) = dsBrokerCN.Tables(0).Rows(0).Item("brokerCode") ' BROKER
                            'End If
                            creditNoteArray(y, 4) = dsCreditNote.Tables(0).Rows(i).Item("creditNoteNumber")
                            creditNoteArray(y, 5) = dsCreditNote.Tables(0).Rows(i).Item("creditNoteDate")
                            creditNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                            creditNoteArray(y, 7) = dsCreditNote.Tables(0).Rows(i).Item("dueDate2")
                            creditNoteArray(y, 8) = dsBrokerCN.Tables(0).Rows(0).Item("brokerageCurrencyCode")
                            creditNoteArray(y, 9) = totalBrokerCN.ToString()
                            creditNoteArray(y, 10) = "0" 'VAT Number
                            creditNoteArray(y, 11) = "2" 'VAT Category
                            creditNoteArray(y, 12) = "0"
                            creditNoteArray(y, 13) = "2.1" 'INC_EXP_CATEGORY
                            creditNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                            creditNoteArray(y, 15) = dsCreditNote.Tables(0).Rows(i).Item("creditNoteDate")
                            creditNoteArray(y, 16) = danaosUsername
                            creditNoteArray(y, 17) = "C" 'DEBIT CREDI
                            creditNoteArray(y, 18) = "" 'VESSEL_CODE
                            Try
                                creditNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                            Catch ex As Exception
                                creditNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("refNum")
                            End Try
                            creditNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                            creditNoteArray(y, 21) = "40" 'JOURNALSERIES
                            creditNoteArray(y, 22) = "0" ' BOOK VALUE
                            creditNoteArray(y, 23) = "0" ' ENTRYRATE
                            creditNoteArray(y, 24) = "True" ' ENTRYRATE

                            ' Adds a Next Row of the Amount + VAT
                            y += 1
                            'Inserts the First Row of the Amount without VAT
                            'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                            creditNoteArray(y, 0) = dsCreditNote.Tables(0).Rows(i).Item("deliveryMergedID")
                            creditNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                            creditNoteArray(y, 2) = "74509"
                            'If i = 1 Then
                            '    creditNoteArray(y, 3) = dsBrokerCN.Tables(0).Rows(1).Item("brokerCode") ' BROKER
                            'ElseIf i = 0 Then
                            creditNoteArray(y, 3) = dsBrokerCN.Tables(0).Rows(0).Item("brokerCode") ' BROKER
                            'End If
                            creditNoteArray(y, 4) = dsCreditNote.Tables(0).Rows(i).Item("creditNoteNumber")
                            creditNoteArray(y, 5) = dsCreditNote.Tables(0).Rows(i).Item("creditNoteDate")
                            creditNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                            creditNoteArray(y, 7) = dsCreditNote.Tables(0).Rows(i).Item("dueDate2")
                            creditNoteArray(y, 8) = dsBrokerCN.Tables(0).Rows(0).Item("brokerageCurrencyCode")
                            creditNoteArray(y, 9) = totalBrokerCN.ToString()
                            creditNoteArray(y, 10) = "0" 'VAT Number
                            creditNoteArray(y, 11) = "2" 'VAT Category
                            creditNoteArray(y, 12) = "0"
                            creditNoteArray(y, 13) = "" 'INC_EXP_CATEGORY
                            creditNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                            creditNoteArray(y, 15) = dsCreditNote.Tables(0).Rows(i).Item("creditNoteDate")
                            creditNoteArray(y, 16) = danaosUsername
                            creditNoteArray(y, 17) = "D" 'DEBIT CREDI
                            creditNoteArray(y, 18) = "" 'VESSEL_CODE
                            Try
                                creditNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                            Catch ex As Exception
                                creditNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("refNum")
                            End Try
                            creditNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                            creditNoteArray(y, 21) = "40" 'JOURNALSERIES
                            creditNoteArray(y, 22) = "0" ' BOOK VALUE
                            creditNoteArray(y, 23) = "0" ' ENTRYRATE
                            creditNoteArray(y, 24) = "True" ' ENTRYRATE
                        Else
                            MessageBox.Show("Credit Note was not exported because the specific Broker Code doesn't exist in Danaos.", "Important", MessageBoxButtons.OK)
                        End If

                    ElseIf dsCreditNote.Tables(0).Rows(i).Item("creditNoteType_ID") = 3 Or dsCreditNote.Tables(0).Rows(i).Item("creditNoteType_ID") = 4 Then

                        'get the merged value find the customer CN items
                        mergedID = dsCreditNote.Tables(0).Rows(i).Item("deliveryMergedID")
                        dsSupplierCN = generaldataset.creditNoteSupplier(orderID, mergedID)
                        Dim totalSupplierCN = 0.0
                        Dim multi = 0.0

                        For g = 0 To dsSupplierCN.Tables(0).Rows.Count - 1
                            totalSupplierCN = totalSupplierCN + dsSupplierCN.Tables(0).Rows(g).Item("creditNoteSupLSAmount")
                        Next

                        ' Get the output values
                        Dim dtData = generaldataset.loadOutputValues(orderID, mergedID)

                        ' Set NA to Not Specified
                        If dtData.Tables("outputvalues").Rows(0).Item("vesselName") = "NA" Then
                            dtData.Tables("outputvalues").Rows(0).Item("vesselName") = "NOT SPECIFIED"
                        End If

                        ' Narrative
                        Dim narrative = dtData.Tables("outputvalues").Rows(0).Item("narativeOfEntry") + "-" + dtData.Tables("outputvalues").Rows(0).Item("vesselName") + "-CREDIT NOTE"

                        Dim exp = generaldataset.exportedInvCN(dsSupplierCN.Tables(0).Rows(0).Item("deliveryID"))

                        If y > 0 Then y += 1

                        'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                        creditNoteArray(y, 0) = dsCreditNote.Tables(0).Rows(i).Item("deliveryMergedID")
                        creditNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                        If dsSupplierCN.Tables(0).Rows(0).Item("creditNoteSupNumber") <> "" Then
                            creditNoteArray(y, 2) = dtData.Tables("outputvalues").Rows(0).Item("supplierCode")
                            creditNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("supplierCode")
                            creditNoteArray(y, 4) = dsSupplierCN.Tables(0).Rows(0).Item("creditNoteSupNumber")
                        ElseIf dsSupplierCN.Tables(0).Rows(0).Item("cnForwNumber") <> "" Then
                            creditNoteArray(y, 2) = dtData.Tables("outputvalues").Rows(0).Item("forwarderCode")
                            creditNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("forwarderCode")
                            creditNoteArray(y, 4) = dsSupplierCN.Tables(0).Rows(0).Item("cnForwNumber")
                        ElseIf dsSupplierCN.Tables(0).Rows(0).Item("cnGenSup1Number") <> "" Then
                            creditNoteArray(y, 2) = dtData.Tables("outputvalues").Rows(0).Item("generalSupplierCode")
                            creditNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("generalSupplierCode")
                            creditNoteArray(y, 4) = dsSupplierCN.Tables(0).Rows(0).Item("cnGenSupExtChar1Price")
                        End If
                        creditNoteArray(y, 5) = dsCreditNote.Tables(0).Rows(i).Item("creditNoteDate")
                        creditNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                        creditNoteArray(y, 7) = dsCreditNote.Tables(0).Rows(i).Item("dueDate2")
                        If totalSupplierCN <> 0 Then
                            creditNoteArray(y, 8) = dtData.Tables("outputvalues").Rows(0).Item("supplierCurrencyCode")
                        ElseIf dsSupplierCN.Tables(0).Rows(0).Item("cnForwFreightPrice") <> 0 Then
                            creditNoteArray(y, 8) = dsSupplierCN.Tables(0).Rows(0).Item("freightBuyingCurr")
                        ElseIf dsSupplierCN.Tables(0).Rows(0).Item("cnGenSupExtChar1Price") <> 0 Then
                            creditNoteArray(y, 8) = dsSupplierCN.Tables(0).Rows(0).Item("extraCharges1BuyingCurr")
                        End If
                        If totalSupplierCN <> 0 Then
                            creditNoteArray(y, 9) = totalSupplierCN.ToString()
                        ElseIf dsSupplierCN.Tables(0).Rows(0).Item("cnForwFreightPrice") <> 0 Then
                            creditNoteArray(y, 9) = dsSupplierCN.Tables(0).Rows(0).Item("cnForwFreightPrice")
                        ElseIf dsSupplierCN.Tables(0).Rows(0).Item("cnGenSupExtChar1Price") <> 0 Then
                            creditNoteArray(y, 9) = dsSupplierCN.Tables(0).Rows(0).Item("cnGenSupExtChar1Price")
                        End If
                        creditNoteArray(y, 10) = "0" 'VAT Number
                        creditNoteArray(y, 11) = "2" 'VAT Category
                        creditNoteArray(y, 12) = "0"
                        creditNoteArray(y, 13) = "2.1" 'INC_EXP_CATEGORY
                        creditNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                        If exp.Tables(0).Rows(0).Item("invExported") = 0 Then
                            creditNoteArray(y, 15) = dtData.Tables("outputvalues").Rows(0).Item("invoiceDate")
                        Else
                            creditNoteArray(y, 15) = dsCreditNote.Tables(0).Rows(i).Item("creditNoteDate")
                        End If
                        creditNoteArray(y, 16) = danaosUsername
                        creditNoteArray(y, 17) = "D" 'DEBIT CREDI
                        creditNoteArray(y, 18) = "" 'VESSEL_CODE
                        Try
                            creditNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                        Catch ex As Exception
                            creditNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("refNum")
                        End Try
                        creditNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                        creditNoteArray(y, 21) = "40" 'JOURNALSERIES
                        creditNoteArray(y, 22) = "0" ' BOOK VALUE
                        creditNoteArray(y, 23) = "0" ' ENTRYRATE
                        creditNoteArray(y, 24) = "True" ' ENTRYRATE

                        ' Adds a Next Row of the Amount + VAT
                        y += 1
                        'Inserts the First Row of the Amount without VAT
                        'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                        creditNoteArray(y, 0) = dsCreditNote.Tables(0).Rows(i).Item("deliveryMergedID")
                        creditNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                        creditNoteArray(y, 2) = dtData.Tables("outputvalues").Rows(0).Item("purchaseCode")
                        If dsSupplierCN.Tables(0).Rows(0).Item("creditNoteSupNumber") <> "" Then
                            creditNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("supplierCode")
                            creditNoteArray(y, 4) = dsSupplierCN.Tables(0).Rows(0).Item("creditNoteSupNumber")
                        ElseIf dsSupplierCN.Tables(0).Rows(0).Item("cnForwNumber") <> "" Then
                            creditNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("forwarderCode")
                            creditNoteArray(y, 4) = dsSupplierCN.Tables(0).Rows(0).Item("cnForwNumber")
                        ElseIf dsSupplierCN.Tables(0).Rows(0).Item("cnGenSup1Number") <> "" Then
                            creditNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("generalSupplierCode")
                            creditNoteArray(y, 4) = dsSupplierCN.Tables(0).Rows(0).Item("cnGenSupExtChar1Price")
                        End If
                        creditNoteArray(y, 5) = dsCreditNote.Tables(0).Rows(i).Item("creditNoteDate")
                        creditNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                        creditNoteArray(y, 7) = dsCreditNote.Tables(0).Rows(i).Item("dueDate2")
                        If totalSupplierCN <> 0 Then
                            creditNoteArray(y, 8) = dtData.Tables("outputvalues").Rows(0).Item("supplierCurrencyCode")
                        ElseIf dsSupplierCN.Tables(0).Rows(0).Item("cnForwFreightPrice") <> 0 Then
                            creditNoteArray(y, 8) = dsSupplierCN.Tables(0).Rows(0).Item("freightBuyingCurr")
                        ElseIf dsSupplierCN.Tables(0).Rows(0).Item("cnGenSupExtChar1Price") <> 0 Then
                            creditNoteArray(y, 8) = dsSupplierCN.Tables(0).Rows(0).Item("extraCharges1BuyingCurr")
                        End If
                        If totalSupplierCN <> 0 Then
                            creditNoteArray(y, 9) = totalSupplierCN.ToString()
                        ElseIf dsSupplierCN.Tables(0).Rows(0).Item("cnForwFreightPrice") <> 0 Then
                            creditNoteArray(y, 9) = dsSupplierCN.Tables(0).Rows(0).Item("cnForwFreightPrice")
                        ElseIf dsSupplierCN.Tables(0).Rows(0).Item("cnGenSupExtChar1Price") <> 0 Then
                            creditNoteArray(y, 9) = dsSupplierCN.Tables(0).Rows(0).Item("cnGenSupExtChar1Price")
                        End If
                        creditNoteArray(y, 10) = "0" 'VAT Number
                        creditNoteArray(y, 11) = "2" 'VAT Category
                        creditNoteArray(y, 12) = "0"
                        creditNoteArray(y, 13) = "" 'INC_EXP_CATEGORY
                        creditNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                        If exp.Tables(0).Rows(0).Item("invExported") = 0 Then
                            creditNoteArray(y, 15) = dtData.Tables("outputvalues").Rows(0).Item("invoiceDate")
                        Else
                            creditNoteArray(y, 15) = dsCreditNote.Tables(0).Rows(i).Item("creditNoteDate")
                        End If
                        creditNoteArray(y, 16) = danaosUsername
                        creditNoteArray(y, 17) = "C" 'DEBIT CREDI
                        creditNoteArray(y, 18) = "" 'VESSEL_CODE
                        Try
                            creditNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                        Catch ex As Exception
                            creditNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("refNum")
                        End Try
                        creditNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                        creditNoteArray(y, 21) = "40" 'JOURNALSERIES
                        creditNoteArray(y, 22) = "0" ' BOOK VALUE
                        creditNoteArray(y, 23) = "0" ' ENTRYRATE
                        creditNoteArray(y, 24) = "True" ' ENTRYRATE
                        ' If CN is Forw (Type 4)
                    End If
                End If
            Next
        End If

        Return creditNoteArray
        'END CREDIT NOTE
        '======================================================================================================================'
    End Function

    ''' <summary>
    ''' DEBIT NOTE
    ''' </summary>
    ''' <returns></returns>
    Public Function debitNote() As String(,)
        '======================================================================================================================'
        ' DEBIT NOTE START
        ' Send the OrderID to get the Credit Note
        Dim dsDebitNote = generaldataset.debitNoteType(orderID)
        Dim rowsCount = dsDebitNote.Tables(0).Rows.Count - 1
        Dim invRowsCount = (rowsCount + 1) * 6
        Dim debitNoteArray(invRowsCount, 24) As String
        Dim dsCustomerDN As DataSet
        Dim dsBrokerDN As DataSet
        Dim y As Integer = 0
        Dim mergedID = 0
        Dim debitNoteID = ""

        ' If has rows
        If dsDebitNote.Tables(0).Rows.Count > 0 Then

            ' Loop on rows
            For i = 0 To dsDebitNote.Tables(0).Rows.Count - 1

                If dsDebitNote.Tables(0).Rows(i).Item("exported") = 0 Then

                    debitNoteID = dsDebitNote.Tables(0).Rows(i).Item("debitNoteID")

                    ' If Debit is Broker
                    If dsDebitNote.Tables(0).Rows(i).Item("debitNoteType_ID") = 1 Then

                        ' Get mergedID of debit and call the dataset
                        Dim totalBrokerDN = 0.0
                        mergedID = dsDebitNote.Tables(0).Rows(i).Item("deliveryMergedID")
                        dsBrokerDN = generaldataset.debitNoteBroker(orderID, mergedID)

                        ' Loop on Dataset to calculate the debit of Broker
                        For o = 0 To dsBrokerDN.Tables("debitnote").Rows.Count - 1
                            totalBrokerDN = totalBrokerDN + dsBrokerDN.Tables(0).Rows(o).Item("debitNoteBrokerLS")
                        Next

                        Dim dtData = generaldataset.loadOutputValues(orderID, mergedID)

                        If dtData.Tables("outputvalues").Rows(0).Item("vesselName") = "NA" Then
                            dtData.Tables("outputvalues").Rows(0).Item("vesselName") = "NOT SPECIFIED"
                        End If

                        Dim narrative = dtData.Tables("outputvalues").Rows(0).Item("narativeOfEntry") + "-" + dtData.Tables("outputvalues").Rows(0).Item("vesselName") + "-DEBIT NOTE"

                        If y > 0 Then y += 1

                        'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                        debitNoteArray(y, 0) = dsDebitNote.Tables(0).Rows(i).Item("deliveryMergedID")
                        debitNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                        debitNoteArray(y, 2) = dsBrokerDN.Tables(0).Rows(0).Item("brokerCode") 'BROKER 
                        debitNoteArray(y, 3) = dsBrokerDN.Tables(0).Rows(0).Item("brokerCode") ' BROKER
                        debitNoteArray(y, 4) = dsBrokerDN.Tables(0).Rows(0).Item("debitNoteNumber")
                        debitNoteArray(y, 5) = dsBrokerDN.Tables(0).Rows(0).Item("debitNoteDate")
                        debitNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                        debitNoteArray(y, 7) = dtData.Tables("outputvalues").Rows(0).Item("dueDate")
                        debitNoteArray(y, 8) = dsBrokerDN.Tables(0).Rows(0).Item("brokerageCurrencyCode")
                        debitNoteArray(y, 9) = totalBrokerDN.ToString()
                        debitNoteArray(y, 10) = "0" 'VAT Number
                        debitNoteArray(y, 11) = "2" 'VAT Category
                        debitNoteArray(y, 12) = "0"
                        debitNoteArray(y, 13) = "2.1" 'INC_EXP_CATEGORY
                        debitNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                        debitNoteArray(y, 15) = dsBrokerDN.Tables(0).Rows(0).Item("debitNoteDate")
                        debitNoteArray(y, 16) = danaosUsername
                        debitNoteArray(y, 17) = "D" 'DEBIT CREDI
                        debitNoteArray(y, 18) = "" 'VESSEL_CODE
                        debitNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                        debitNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                        debitNoteArray(y, 21) = "40" 'JOURNALSERIES
                        debitNoteArray(y, 22) = "0" ' BOOK VALUE
                        debitNoteArray(y, 23) = "0" ' ENTRYRATE
                        debitNoteArray(y, 24) = "True" ' ENTRYRATE
                        y += 1

                        'Inserts the First Row of the Amount without VAT
                        'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                        debitNoteArray(y, 0) = dsDebitNote.Tables(0).Rows(i).Item("deliveryMergedID")
                        debitNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                        debitNoteArray(y, 2) = "74509"
                        debitNoteArray(y, 3) = dsBrokerDN.Tables(0).Rows(0).Item("brokerCode") '' BROKER CODE
                        debitNoteArray(y, 4) = dsBrokerDN.Tables(0).Rows(0).Item("debitNoteNumber")
                        debitNoteArray(y, 5) = dsBrokerDN.Tables(0).Rows(0).Item("debitNoteDate")
                        debitNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                        debitNoteArray(y, 7) = dtData.Tables("outputvalues").Rows(0).Item("dueDate")
                        debitNoteArray(y, 8) = dsBrokerDN.Tables(0).Rows(0).Item("brokerageCurrencyCode")
                        debitNoteArray(y, 9) = totalBrokerDN.ToString()
                        debitNoteArray(y, 10) = "0" 'VAT Number
                        debitNoteArray(y, 11) = "2" 'VAT Category
                        debitNoteArray(y, 12) = "0"
                        debitNoteArray(y, 13) = "" 'INC_EXP_CATEGORY
                        debitNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                        debitNoteArray(y, 15) = dsBrokerDN.Tables(0).Rows(0).Item("debitNoteDate")
                        debitNoteArray(y, 16) = danaosUsername
                        debitNoteArray(y, 17) = "C" 'DEBIT CREDI
                        debitNoteArray(y, 18) = "" 'VESSEL_CODE
                        debitNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                        debitNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                        debitNoteArray(y, 21) = "40" 'JOURNALSERIES
                        debitNoteArray(y, 22) = "0" ' BOOK VALUE
                        debitNoteArray(y, 23) = "0" ' ENTRYRATE
                        debitNoteArray(y, 24) = "True" ' ENTRYRATE

                        'Adds a Next Row of the Amount + VAT


                        ' If debit is Customer
                    ElseIf dsDebitNote.Tables(0).Rows(i).Item("debitNoteType_ID") = 2 Then

                        Dim multiBuying = 0.0
                        Dim multiSelling = 0.0
                        Dim totalItemsSelling = 0.0
                        Dim totalItemsBuying = 0.0
                        Dim sellVat = 0.0
                        Dim buyVat = 0.0
                        Dim totalCustomerDNBuying = 0.0
                        Dim totalCustomerDNSelling = 0.0

                        ' Get the values with the vatids
                        mergedID = dsDebitNote.Tables(0).Rows(i).Item("deliveryMergedID")
                        dsCustomerDN = generaldataset.debitNoteCustomer(orderID, mergedID)

                        ' loop on values and calculate vat then add the vat to final value
                        For o = 0 To dsCustomerDN.Tables(0).Rows.Count - 1
                            ' Calculate Items Buying value
                            multiBuying = dsCustomerDN.Tables(0).Rows(o).Item("debitNoteItemBuyingValue") * dsCustomerDN.Tables(0).Rows(o).Item("debitNoteItemQuantity")
                            totalItemsBuying = totalItemsBuying + Math.Round(multiBuying, 2, MidpointRounding.AwayFromZero)

                            ' Calculate Items Selling value
                            multiSelling = dsCustomerDN.Tables(0).Rows(o).Item("debitNoteItemSellingValue") * dsCustomerDN.Tables(0).Rows(o).Item("debitNoteItemQuantity")
                            totalItemsSelling = totalItemsSelling + Math.Round(multiSelling, 2, MidpointRounding.AwayFromZero)

                            ' calculate vat based on ID
                            buyVat = buyVat + calculateVat(totalItemsBuying, dsCustomerDN.Tables(0).Rows(o).Item("debitNoteItemBuyingVat_ID"))
                            sellVat = sellVat + calculateVat(totalItemsSelling, dsCustomerDN.Tables(0).Rows(o).Item("debitNoteItemSellingVat_ID"))

                            ' add vat to total price
                            totalCustomerDNBuying = totalItemsBuying + buyVat
                            totalCustomerDNSelling = totalItemsSelling + sellVat
                        Next

                        ' Get data for the output 
                        Dim dtData = generaldataset.loadOutputValues(orderID, mergedID)

                        If dtData.Tables("outputvalues").Rows(0).Item("vesselName") = "NA" Then
                            dtData.Tables("outputvalues").Rows(0).Item("vesselName") = "NOT SPECIFIED"
                        End If

                        Dim narrative = dtData.Tables("outputvalues").Rows(0).Item("narativeOfEntry") + "-" + dtData.Tables("outputvalues").Rows(0).Item("vesselName") + "-DEBIT NOTE CUSTOMER"

                        If multiSelling > 0 Then

                            ' if vat selling greater than 0 then print 3 lines else print 2 
                            If sellVat > 0 Then

                                If y > 0 Then y += 1

                                'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                                debitNoteArray(y, 0) = dsDebitNote.Tables(0).Rows(i).Item("deliveryMergedID")
                                debitNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                                debitNoteArray(y, 2) = dtData.Tables("outputvalues").Rows(0).Item("customerCode")
                                debitNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("customerCode")
                                debitNoteArray(y, 4) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteNumber")
                                debitNoteArray(y, 5) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteDate")
                                debitNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                                debitNoteArray(y, 7) = dsDebitNote.Tables(0).Rows(i).Item("dueDate2")
                                debitNoteArray(y, 8) = dtData.Tables("outputvalues").Rows(0).Item("customerCurrencyCode")
                                debitNoteArray(y, 9) = totalCustomerDNSelling.ToString()
                                debitNoteArray(y, 10) = "0" 'VAT Number
                                debitNoteArray(y, 11) = "7" 'VAT Category
                                debitNoteArray(y, 12) = sellVat.ToString()
                                debitNoteArray(y, 13) = "1.7" 'INC_EXP_CATEGORY
                                debitNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                                debitNoteArray(y, 15) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteCreatedDate")
                                debitNoteArray(y, 16) = danaosUsername
                                debitNoteArray(y, 17) = "D" 'DEBIT CREDI
                                debitNoteArray(y, 18) = "" 'VESSEL_CODE
                                debitNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("refNum")
                                debitNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                                debitNoteArray(y, 21) = "40" 'JOURNALSERIES
                                debitNoteArray(y, 22) = "0" ' BOOK VALUE
                                debitNoteArray(y, 23) = "0" ' ENTRYRATE
                                debitNoteArray(y, 24) = "True"

                                ' Adds a Next Row of the Amount + VAT
                                y += 1
                                'Inserts the First Row of the Amount without VAT
                                'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                                debitNoteArray(y, 0) = dsDebitNote.Tables(0).Rows(i).Item("deliveryMergedID")
                                debitNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                                debitNoteArray(y, 2) = dtData.Tables("outputvalues").Rows(0).Item("salesCode")
                                debitNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("customerCode")
                                debitNoteArray(y, 4) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteNumber")
                                debitNoteArray(y, 5) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteDate")
                                debitNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                                debitNoteArray(y, 7) = dsDebitNote.Tables(0).Rows(i).Item("dueDate2")
                                debitNoteArray(y, 8) = dtData.Tables("outputvalues").Rows(0).Item("customerCurrencyCode")
                                debitNoteArray(y, 9) = totalItemsSelling.ToString()
                                debitNoteArray(y, 10) = "0" 'VAT Number
                                debitNoteArray(y, 11) = "2" 'VAT Category
                                debitNoteArray(y, 12) = "0"
                                debitNoteArray(y, 13) = "" 'INC_EXP_CATEGORY
                                debitNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                                debitNoteArray(y, 15) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteCreatedDate")
                                debitNoteArray(y, 16) = danaosUsername
                                debitNoteArray(y, 17) = "C" 'DEBIT CREDI
                                debitNoteArray(y, 18) = "" 'VESSEL_CODE
                                debitNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                                debitNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                                debitNoteArray(y, 21) = "40" 'JOURNALSERIES
                                debitNoteArray(y, 22) = "0" ' BOOK VALUE
                                debitNoteArray(y, 23) = "0" ' ENTRYRATE
                                debitNoteArray(y, 24) = "True"

                                'Adds a THird Row of the VAT AMOunt only
                                y += 1
                                'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                                debitNoteArray(y, 0) = dsDebitNote.Tables(0).Rows(i).Item("deliveryMergedID")
                                debitNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                                debitNoteArray(y, 2) = "38602"
                                debitNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("customerCode")
                                debitNoteArray(y, 4) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteNumber")
                                debitNoteArray(y, 5) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteDate")
                                debitNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                                debitNoteArray(y, 7) = dsDebitNote.Tables(0).Rows(i).Item("dueDate2")
                                debitNoteArray(y, 8) = dtData.Tables("outputvalues").Rows(0).Item("customerCurrencyCode")
                                debitNoteArray(y, 9) = sellVat.ToString()
                                debitNoteArray(y, 10) = "0" 'VAT Number
                                debitNoteArray(y, 11) = "2" 'VAT Category
                                debitNoteArray(y, 12) = "0"
                                debitNoteArray(y, 13) = "" 'INC_EXP_CATEGORY
                                debitNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                                debitNoteArray(y, 15) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteCreatedDate")
                                debitNoteArray(y, 16) = danaosUsername
                                debitNoteArray(y, 17) = "C" 'DEBIT CREDI
                                debitNoteArray(y, 18) = "" 'VESSEL_CODE
                                debitNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                                debitNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                                debitNoteArray(y, 21) = "40" 'JOURNALSERIES
                                debitNoteArray(y, 22) = "0" ' BOOK VALUE
                                debitNoteArray(y, 23) = "0" ' ENTRYRATE
                                debitNoteArray(y, 24) = "True"

                            Else
                                If y > 0 Then y += 1

                                'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                                debitNoteArray(y, 0) = dsDebitNote.Tables(0).Rows(i).Item("deliveryMergedID")
                                debitNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                                debitNoteArray(y, 2) = dtData.Tables("outputvalues").Rows(0).Item("customerCode")
                                debitNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("customerCode")
                                debitNoteArray(y, 4) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteNumber")
                                debitNoteArray(y, 5) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteDate")
                                debitNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                                debitNoteArray(y, 7) = dsDebitNote.Tables(0).Rows(i).Item("dueDate2")
                                debitNoteArray(y, 8) = dtData.Tables("outputvalues").Rows(0).Item("customerCurrencyCode")
                                debitNoteArray(y, 9) = totalCustomerDNSelling.ToString()
                                debitNoteArray(y, 10) = "0" 'VAT Number
                                debitNoteArray(y, 11) = "2" 'VAT Category
                                debitNoteArray(y, 12) = sellVat.ToString()
                                debitNoteArray(y, 13) = "1.1" 'INC_EXP_CATEGORY
                                debitNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                                debitNoteArray(y, 15) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteCreatedDate")
                                debitNoteArray(y, 16) = danaosUsername
                                debitNoteArray(y, 17) = "D" 'DEBIT CREDI
                                debitNoteArray(y, 18) = "" 'VESSEL_CODE
                                debitNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                                debitNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                                debitNoteArray(y, 21) = "40" 'JOURNALSERIES
                                debitNoteArray(y, 22) = "0" ' BOOK VALUE
                                debitNoteArray(y, 23) = "0" ' ENTRYRATE
                                debitNoteArray(y, 24) = "True"

                                ' Adds a Next Row of the Amount + VAT
                                y += 1
                                'Inserts the First Row of the Amount without VAT
                                'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                                debitNoteArray(y, 0) = dsDebitNote.Tables(0).Rows(i).Item("deliveryMergedID")
                                debitNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                                debitNoteArray(y, 2) = dtData.Tables("outputvalues").Rows(0).Item("salesCode")
                                debitNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("customerCode")
                                debitNoteArray(y, 4) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteNumber")
                                debitNoteArray(y, 5) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteDate")
                                debitNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                                debitNoteArray(y, 7) = dsDebitNote.Tables(0).Rows(i).Item("dueDate2")
                                debitNoteArray(y, 8) = dtData.Tables("outputvalues").Rows(0).Item("customerCurrencyCode")
                                debitNoteArray(y, 9) = totalCustomerDNSelling.ToString()
                                debitNoteArray(y, 10) = "0" 'VAT Number
                                debitNoteArray(y, 11) = "2" 'VAT Category
                                debitNoteArray(y, 12) = sellVat.ToString()
                                debitNoteArray(y, 13) = "" 'INC_EXP_CATEGORY
                                debitNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                                debitNoteArray(y, 15) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteCreatedDate")
                                debitNoteArray(y, 16) = danaosUsername
                                debitNoteArray(y, 17) = "C" 'DEBIT CREDI
                                debitNoteArray(y, 18) = "" 'VESSEL_CODE
                                debitNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                                debitNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                                debitNoteArray(y, 21) = "40" 'JOURNALSERIES
                                debitNoteArray(y, 22) = "0" ' BOOK VALUE
                                debitNoteArray(y, 23) = "0" ' ENTRYRATE
                                debitNoteArray(y, 24) = "True"
                            End If
                        End If

                        narrative = dtData.Tables("outputvalues").Rows(0).Item("narativeOfEntry") + "-" + dtData.Tables("outputvalues").Rows(0).Item("vesselName") + "- DEBIT NOTE SUPPLIER"

                        If multiBuying > 0 Then
                            ' if vat buying greater than 0 then print 3 lines else print 2
                            If buyVat > 0 Then
                                If y > 0 Then y += 1

                                'invoicesBuyArray(u, 0) = dsBuyingInvoices.Tables("invoicesSelling").Rows(y).Item("deliveryID")
                                debitNoteArray(y, 0) = dsDebitNote.Tables(0).Rows(i).Item("deliveryMergedID")
                                debitNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                                If dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierType_ID") = 1 Then
                                    debitNoteArray(y, 2) = dtData.Tables("outputvalues").Rows(0).Item("supplierCode")
                                    debitNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("supplierCode")
                                ElseIf dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierType_ID") = 2 Then
                                    debitNoteArray(y, 2) = dsCustomerDN.Tables(0).Rows(i).Item("forwarderCode")
                                    debitNoteArray(y, 3) = dsCustomerDN.Tables(0).Rows(i).Item("forwarderCode")
                                ElseIf dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierType_ID") = 3 Then
                                    debitNoteArray(y, 2) = dsCustomerDN.Tables(0).Rows(i).Item("clearingAgentCode")
                                    debitNoteArray(y, 3) = dsCustomerDN.Tables(0).Rows(i).Item("clearingAgentCode")
                                End If
                                debitNoteArray(y, 4) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierInvoiceNumber")
                                debitNoteArray(y, 5) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierInvoiceDate")
                                debitNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                                debitNoteArray(y, 7) = dsDebitNote.Tables(0).Rows(i).Item("dueDate2")
                                debitNoteArray(y, 8) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteSellingCurrencyCode")
                                debitNoteArray(y, 9) = totalCustomerDNBuying.ToString()
                                debitNoteArray(y, 10) = "0" 'VAT Number
                                debitNoteArray(y, 11) = "7" 'VAT Category
                                debitNoteArray(y, 12) = buyVat.ToString()
                                debitNoteArray(y, 13) = "2.1" 'INC_EXP_CATEGORY
                                debitNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                                debitNoteArray(y, 15) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteCreatedDate")
                                debitNoteArray(y, 16) = danaosUsername
                                debitNoteArray(y, 17) = "C" 'DEBIT CREDI
                                debitNoteArray(y, 18) = "" 'VESSEL_CODE
                                debitNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                                debitNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                                debitNoteArray(y, 21) = "40" 'JOURNALSERIES
                                debitNoteArray(y, 22) = "0" ' BOOK VALUE
                                debitNoteArray(y, 23) = "0" ' ENTRYRATE
                                debitNoteArray(y, 24) = "True"

                                ' insert total + vat
                                y += 1
                                'Inserts the First Row of the Amount without VAT
                                'invoicesBuyArray(u, 0) = dsBuyingInvoices.Tables("invoicesSelling").Rows(y).Item("deliveryID")
                                debitNoteArray(y, 0) = dsDebitNote.Tables(0).Rows(i).Item("deliveryMergedID")
                                debitNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                                debitNoteArray(y, 2) = dtData.Tables("outputvalues").Rows(0).Item("purchaseCode")
                                If dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierType_ID") = 1 Then
                                    debitNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("supplierCode")
                                ElseIf dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierType_ID") = 2 Then
                                    debitNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("forwardedCode")
                                ElseIf dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierType_ID") = 3 Then
                                    debitNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("clearingAgentCode")
                                End If
                                debitNoteArray(y, 4) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierInvoiceNumber")
                                debitNoteArray(y, 5) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierInvoiceDate")
                                debitNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                                debitNoteArray(y, 7) = dsDebitNote.Tables(0).Rows(i).Item("dueDate2")
                                debitNoteArray(y, 8) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteSellingCurrencyCode")
                                debitNoteArray(y, 9) = totalItemsBuying.ToString()
                                debitNoteArray(y, 10) = "0" 'VAT Number
                                debitNoteArray(y, 11) = "2" 'VAT Category
                                debitNoteArray(y, 12) = "0"
                                debitNoteArray(y, 13) = "" 'INC_EXP_CATEGORY
                                debitNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                                debitNoteArray(y, 15) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteCreatedDate")
                                debitNoteArray(y, 16) = danaosUsername
                                debitNoteArray(y, 17) = "D" 'DEBIT CREDI
                                debitNoteArray(y, 18) = "" 'VESSEL_CODE
                                debitNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                                debitNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                                debitNoteArray(y, 21) = "40" 'JOURNALSERIES
                                debitNoteArray(y, 22) = "0" ' BOOK VALUE
                                debitNoteArray(y, 23) = "0" ' ENTRYRATE
                                debitNoteArray(y, 24) = "True"

                                ' if vat is 0 skip dont insert anything
                                y += 1
                                'invoicesBuyArray(u, 0) = dsBuyingInvoices.Tables("invoicesSelling").Rows(y).Item("deliveryID")
                                debitNoteArray(y, 0) = dsDebitNote.Tables(0).Rows(i).Item("deliveryMergedID")
                                debitNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                                debitNoteArray(y, 2) = "38601"
                                debitNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("purchaseCode")
                                debitNoteArray(y, 4) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierInvoiceNumber")
                                debitNoteArray(y, 5) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierInvoiceDate")
                                debitNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                                debitNoteArray(y, 7) = dsDebitNote.Tables(0).Rows(i).Item("dueDate2")
                                debitNoteArray(y, 8) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteSellingCurrencyCode")
                                debitNoteArray(y, 9) = buyVat.ToString()
                                debitNoteArray(y, 10) = "0" 'VAT Number
                                debitNoteArray(y, 11) = "2" 'VAT Category
                                debitNoteArray(y, 12) = "0"
                                debitNoteArray(y, 13) = "" 'INC_EXP_CATEGORY
                                debitNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                                debitNoteArray(y, 15) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteCreatedDate")
                                debitNoteArray(y, 16) = danaosUsername
                                debitNoteArray(y, 17) = "D" 'DEBIT CREDI
                                debitNoteArray(y, 18) = "" 'VESSEL_CODE
                                debitNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                                debitNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                                debitNoteArray(y, 21) = "40" 'JOURNALSERIES
                                debitNoteArray(y, 22) = "0" ' BOOK VALUE                        
                                debitNoteArray(y, 23) = "0" ' ENTRYRATE
                                debitNoteArray(y, 24) = "True"

                            Else
                                If y > 0 Then y += 1

                                'invoicesBuyArray(u, 0) = dsBuyingInvoices.Tables("invoicesSelling").Rows(y).Item("deliveryID")
                                debitNoteArray(y, 0) = dsDebitNote.Tables(0).Rows(i).Item("deliveryMergedID")
                                debitNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                                If dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierType_ID") = 1 Then
                                    debitNoteArray(y, 2) = dtData.Tables("outputvalues").Rows(0).Item("supplierCode")
                                    debitNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("supplierCode")
                                ElseIf dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierType_ID") = 2 Then
                                    debitNoteArray(y, 2) = dsCustomerDN.Tables(0).Rows(i).Item("forwarderCode")
                                    debitNoteArray(y, 3) = dsCustomerDN.Tables(0).Rows(i).Item("forwarderCode")
                                ElseIf dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierType_ID") = 3 Then
                                    debitNoteArray(y, 2) = dsCustomerDN.Tables(0).Rows(i).Item("clearingAgentCode")
                                    debitNoteArray(y, 3) = dsCustomerDN.Tables(0).Rows(i).Item("clearingAgentCode")
                                End If
                                debitNoteArray(y, 4) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierInvoiceNumber")
                                debitNoteArray(y, 5) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierInvoiceDate")
                                debitNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                                debitNoteArray(y, 7) = dsDebitNote.Tables(0).Rows(i).Item("dueDate2")
                                debitNoteArray(y, 8) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteSellingCurrencyCode")
                                debitNoteArray(y, 9) = totalCustomerDNBuying.ToString()
                                debitNoteArray(y, 10) = "0" 'VAT Number
                                debitNoteArray(y, 11) = "2" 'VAT Category
                                debitNoteArray(y, 12) = buyVat.ToString()
                                debitNoteArray(y, 13) = "2.1" 'INC_EXP_CATEGORY
                                debitNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                                debitNoteArray(y, 15) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteCreatedDate")
                                debitNoteArray(y, 16) = danaosUsername
                                debitNoteArray(y, 17) = "C" 'DEBIT CREDI
                                debitNoteArray(y, 18) = "" 'VESSEL_CODE
                                debitNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                                debitNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                                debitNoteArray(y, 21) = "40" 'JOURNALSERIES
                                debitNoteArray(y, 22) = "0" ' BOOK VALUE
                                debitNoteArray(y, 23) = "0" ' ENTRYRATE
                                debitNoteArray(y, 24) = "True"

                                ' insert total + vat
                                y += 1
                                'Inserts the First Row of the Amount without VAT
                                'invoicesBuyArray(u, 0) = dsBuyingInvoices.Tables("invoicesSelling").Rows(y).Item("deliveryID")
                                debitNoteArray(y, 0) = dsDebitNote.Tables(0).Rows(i).Item("deliveryMergedID")
                                debitNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                                debitNoteArray(y, 2) = dtData.Tables("outputvalues").Rows(0).Item("purchaseCode")
                                If dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierType_ID") = 1 Then
                                    debitNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("supplierCode")
                                ElseIf dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierType_ID") = 2 Then
                                    debitNoteArray(y, 3) = dsCustomerDN.Tables(0).Rows(i).Item("forwarderCode")
                                ElseIf dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierType_ID") = 3 Then
                                    debitNoteArray(y, 3) = dsCustomerDN.Tables(0).Rows(i).Item("clearingAgentCode")
                                End If
                                debitNoteArray(y, 4) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierInvoiceNumber")
                                debitNoteArray(y, 5) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierInvoiceDate")
                                debitNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                                debitNoteArray(y, 7) = dsDebitNote.Tables(0).Rows(i).Item("dueDate2")
                                debitNoteArray(y, 8) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteSellingCurrencyCode")
                                debitNoteArray(y, 9) = totalCustomerDNBuying.ToString()
                                debitNoteArray(y, 10) = "0" 'VAT Number
                                debitNoteArray(y, 11) = "2" 'VAT Category
                                debitNoteArray(y, 12) = buyVat.ToString()
                                debitNoteArray(y, 13) = "" 'INC_EXP_CATEGORY
                                debitNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                                debitNoteArray(y, 15) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteCreatedDate")
                                debitNoteArray(y, 16) = danaosUsername
                                debitNoteArray(y, 17) = "D" 'DEBIT CREDI
                                debitNoteArray(y, 18) = "" 'VESSEL_CODE
                                debitNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                                debitNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                                debitNoteArray(y, 21) = "40" 'JOURNALSERIES
                                debitNoteArray(y, 22) = "0" ' BOOK VALUE
                                debitNoteArray(y, 23) = "0" ' ENTRYRATE
                                debitNoteArray(y, 24) = "True"
                            End If
                        End If
                    ElseIf dsDebitNote.Tables(0).Rows(i).Item("debitNoteType_ID") = 3 Then
                        Dim multiBuying = 0.0
                        Dim multiSelling = 0.0
                        Dim totalItemsSelling = 0.0
                        Dim totalItemsBuying = 0.0
                        Dim sellVat = 0.0
                        Dim buyVat = 0.0
                        Dim totalCustomerDNBuying = 0.0
                        Dim totalCustomerDNSelling = 0.0

                        ' Get the values with the vatids
                        mergedID = dsDebitNote.Tables(0).Rows(i).Item("deliveryMergedID")
                        dsCustomerDN = generaldataset.debitNoteCustomer(orderID, mergedID)

                        ' loop on values and calculate vat then add the vat to final value
                        For o = 0 To dsCustomerDN.Tables(0).Rows.Count - 1
                            ' Calculate Items Buying value
                            multiBuying = dsCustomerDN.Tables(0).Rows(o).Item("debitNoteItemBuyingValue") * dsCustomerDN.Tables(0).Rows(o).Item("debitNoteItemQuantity")
                            totalItemsBuying = totalItemsBuying + Math.Round(multiBuying, 2, MidpointRounding.AwayFromZero)

                            ' Calculate Items Selling value
                            multiSelling = dsCustomerDN.Tables(0).Rows(o).Item("debitNoteItemSellingValue") * dsCustomerDN.Tables(0).Rows(o).Item("debitNoteItemQuantity")
                            totalItemsSelling = totalItemsSelling + Math.Round(multiSelling, 2, MidpointRounding.AwayFromZero)

                            ' calculate vat based on ID
                            buyVat = buyVat + calculateVat(totalItemsBuying, dsCustomerDN.Tables(0).Rows(o).Item("debitNoteItemBuyingVat_ID"))
                            sellVat = sellVat + calculateVat(totalItemsSelling, dsCustomerDN.Tables(0).Rows(o).Item("debitNoteItemSellingVat_ID"))

                            ' add vat to total price
                            totalCustomerDNBuying = totalItemsBuying + buyVat
                            totalCustomerDNSelling = totalItemsSelling + sellVat
                        Next

                        ' Get data for the output 
                        Dim dtData = generaldataset.loadOutputValues(orderID, mergedID)

                        If dtData.Tables("outputvalues").Rows(0).Item("vesselName") = "NA" Then
                            dtData.Tables("outputvalues").Rows(0).Item("vesselName") = "NOT SPECIFIED"
                        End If

                        Dim narrative = dtData.Tables("outputvalues").Rows(0).Item("narativeOfEntry") + "-" + dtData.Tables("outputvalues").Rows(0).Item("vesselName") + "-DEBIT NOTE CUSTOMER"

                        If multiSelling > 0 Then

                            ' if vat selling greater than 0 then print 3 lines else print 2 
                            If sellVat > 0 Then

                                If y > 0 Then y += 1

                                'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                                debitNoteArray(y, 0) = dsDebitNote.Tables(0).Rows(i).Item("deliveryMergedID")
                                debitNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                                debitNoteArray(y, 2) = dtData.Tables("outputvalues").Rows(0).Item("customerCode")
                                debitNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("customerCode")
                                debitNoteArray(y, 4) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteNumber")
                                debitNoteArray(y, 5) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteDate")
                                debitNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                                debitNoteArray(y, 7) = dsDebitNote.Tables(0).Rows(i).Item("dueDate2")
                                debitNoteArray(y, 8) = dtData.Tables("outputvalues").Rows(0).Item("customerCurrencyCode")
                                debitNoteArray(y, 9) = totalCustomerDNSelling.ToString()
                                debitNoteArray(y, 10) = "0" 'VAT Number
                                debitNoteArray(y, 11) = "7" 'VAT Category
                                debitNoteArray(y, 12) = sellVat.ToString()
                                debitNoteArray(y, 13) = "1.7" 'INC_EXP_CATEGORY
                                debitNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                                debitNoteArray(y, 15) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteCreatedDate")
                                debitNoteArray(y, 16) = danaosUsername
                                debitNoteArray(y, 17) = "C" 'DEBIT CREDI
                                debitNoteArray(y, 18) = "" 'VESSEL_CODE
                                debitNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("refNum")
                                debitNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                                debitNoteArray(y, 21) = "40" 'JOURNALSERIES
                                debitNoteArray(y, 22) = "0" ' BOOK VALUE
                                debitNoteArray(y, 23) = "0" ' ENTRYRATE
                                debitNoteArray(y, 24) = "True"

                                ' Adds a Next Row of the Amount + VAT
                                y += 1
                                'Inserts the First Row of the Amount without VAT
                                'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                                debitNoteArray(y, 0) = dsDebitNote.Tables(0).Rows(i).Item("deliveryMergedID")
                                debitNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                                debitNoteArray(y, 2) = dtData.Tables("outputvalues").Rows(0).Item("salesCode")
                                debitNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("customerCode")
                                debitNoteArray(y, 4) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteNumber")
                                debitNoteArray(y, 5) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteDate")
                                debitNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                                debitNoteArray(y, 7) = dsDebitNote.Tables(0).Rows(i).Item("dueDate2")
                                debitNoteArray(y, 8) = dtData.Tables("outputvalues").Rows(0).Item("customerCurrencyCode")
                                debitNoteArray(y, 9) = totalItemsSelling.ToString()
                                debitNoteArray(y, 10) = "0" 'VAT Number
                                debitNoteArray(y, 11) = "2" 'VAT Category
                                debitNoteArray(y, 12) = "0"
                                debitNoteArray(y, 13) = "" 'INC_EXP_CATEGORY
                                debitNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                                debitNoteArray(y, 15) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteCreatedDate")
                                debitNoteArray(y, 16) = danaosUsername
                                debitNoteArray(y, 17) = "D" 'DEBIT CREDI
                                debitNoteArray(y, 18) = "" 'VESSEL_CODE
                                debitNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                                debitNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                                debitNoteArray(y, 21) = "40" 'JOURNALSERIES
                                debitNoteArray(y, 22) = "0" ' BOOK VALUE
                                debitNoteArray(y, 23) = "0" ' ENTRYRATE
                                debitNoteArray(y, 24) = "True"

                                'Adds a THird Row of the VAT AMOunt only
                                y += 1
                                'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                                debitNoteArray(y, 0) = dsDebitNote.Tables(0).Rows(i).Item("deliveryMergedID")
                                debitNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                                debitNoteArray(y, 2) = "38602"
                                debitNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("customerCode")
                                debitNoteArray(y, 4) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteNumber")
                                debitNoteArray(y, 5) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteDate")
                                debitNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                                debitNoteArray(y, 7) = dsDebitNote.Tables(0).Rows(i).Item("dueDate2")
                                debitNoteArray(y, 8) = dtData.Tables("outputvalues").Rows(0).Item("customerCurrencyCode")
                                debitNoteArray(y, 9) = sellVat.ToString()
                                debitNoteArray(y, 10) = "0" 'VAT Number
                                debitNoteArray(y, 11) = "2" 'VAT Category
                                debitNoteArray(y, 12) = "0"
                                debitNoteArray(y, 13) = "" 'INC_EXP_CATEGORY
                                debitNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                                debitNoteArray(y, 15) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteCreatedDate")
                                debitNoteArray(y, 16) = danaosUsername
                                debitNoteArray(y, 17) = "D" 'DEBIT CREDI
                                debitNoteArray(y, 18) = "" 'VESSEL_CODE
                                debitNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                                debitNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                                debitNoteArray(y, 21) = "40" 'JOURNALSERIES
                                debitNoteArray(y, 22) = "0" ' BOOK VALUE
                                debitNoteArray(y, 23) = "0" ' ENTRYRATE
                                debitNoteArray(y, 24) = "True"

                            Else
                                If y > 0 Then y += 1

                                'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                                debitNoteArray(y, 0) = dsDebitNote.Tables(0).Rows(i).Item("deliveryMergedID")
                                debitNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                                debitNoteArray(y, 2) = dtData.Tables("outputvalues").Rows(0).Item("customerCode")
                                debitNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("customerCode")
                                debitNoteArray(y, 4) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteNumber")
                                debitNoteArray(y, 5) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteDate")
                                debitNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                                debitNoteArray(y, 7) = dsDebitNote.Tables(0).Rows(i).Item("dueDate2")
                                debitNoteArray(y, 8) = dtData.Tables("outputvalues").Rows(0).Item("customerCurrencyCode")
                                debitNoteArray(y, 9) = totalCustomerDNSelling.ToString()
                                debitNoteArray(y, 10) = "0" 'VAT Number
                                debitNoteArray(y, 11) = "2" 'VAT Category
                                debitNoteArray(y, 12) = sellVat.ToString()
                                debitNoteArray(y, 13) = "1.1" 'INC_EXP_CATEGORY
                                debitNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                                debitNoteArray(y, 15) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteCreatedDate")
                                debitNoteArray(y, 16) = danaosUsername
                                debitNoteArray(y, 17) = "C" 'DEBIT CREDI
                                debitNoteArray(y, 18) = "" 'VESSEL_CODE
                                debitNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                                debitNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                                debitNoteArray(y, 21) = "40" 'JOURNALSERIES
                                debitNoteArray(y, 22) = "0" ' BOOK VALUE
                                debitNoteArray(y, 23) = "0" ' ENTRYRATE
                                debitNoteArray(y, 24) = "True"

                                ' Adds a Next Row of the Amount + VAT
                                y += 1
                                'Inserts the First Row of the Amount without VAT
                                'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                                debitNoteArray(y, 0) = dsDebitNote.Tables(0).Rows(i).Item("deliveryMergedID")
                                debitNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                                debitNoteArray(y, 2) = dtData.Tables("outputvalues").Rows(0).Item("salesCode")
                                debitNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("customerCode")
                                debitNoteArray(y, 4) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteNumber")
                                debitNoteArray(y, 5) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteDate")
                                debitNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                                debitNoteArray(y, 7) = dsDebitNote.Tables(0).Rows(i).Item("dueDate2")
                                debitNoteArray(y, 8) = dtData.Tables("outputvalues").Rows(0).Item("customerCurrencyCode")
                                debitNoteArray(y, 9) = totalCustomerDNSelling.ToString()
                                debitNoteArray(y, 10) = "0" 'VAT Number
                                debitNoteArray(y, 11) = "2" 'VAT Category
                                debitNoteArray(y, 12) = sellVat.ToString()
                                debitNoteArray(y, 13) = "" 'INC_EXP_CATEGORY
                                debitNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                                debitNoteArray(y, 15) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteCreatedDate")
                                debitNoteArray(y, 16) = danaosUsername
                                debitNoteArray(y, 17) = "D" 'DEBIT CREDI
                                debitNoteArray(y, 18) = "" 'VESSEL_CODE
                                debitNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                                debitNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                                debitNoteArray(y, 21) = "40" 'JOURNALSERIES
                                debitNoteArray(y, 22) = "0" ' BOOK VALUE
                                debitNoteArray(y, 23) = "0" ' ENTRYRATE
                                debitNoteArray(y, 24) = "True"
                            End If
                        End If

                        narrative = dtData.Tables("outputvalues").Rows(0).Item("narativeOfEntry") + "-" + dtData.Tables("outputvalues").Rows(0).Item("vesselName") + "- DEBIT NOTE SUPPLIER"

                        If multiBuying > 0 Then
                            ' if vat buying greater than 0 then print 3 lines else print 2
                            If buyVat > 0 Then
                                If y > 0 Then y += 1

                                'invoicesBuyArray(u, 0) = dsBuyingInvoices.Tables("invoicesSelling").Rows(y).Item("deliveryID")
                                debitNoteArray(y, 0) = dsDebitNote.Tables(0).Rows(i).Item("deliveryMergedID")
                                debitNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                                If dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierType_ID") = 1 Then
                                    debitNoteArray(y, 2) = dtData.Tables("outputvalues").Rows(0).Item("supplierCode")
                                    debitNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("supplierCode")
                                ElseIf dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierType_ID") = 2 Then
                                    debitNoteArray(y, 2) = dsCustomerDN.Tables(0).Rows(i).Item("forwarderCode")
                                    debitNoteArray(y, 3) = dsCustomerDN.Tables(0).Rows(i).Item("forwarderCode")
                                ElseIf dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierType_ID") = 3 Then
                                    debitNoteArray(y, 2) = dsCustomerDN.Tables(0).Rows(i).Item("clearingAgentCode")
                                    debitNoteArray(y, 3) = dsCustomerDN.Tables(0).Rows(i).Item("clearingAgentCode")
                                End If
                                debitNoteArray(y, 4) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierInvoiceNumber")
                                debitNoteArray(y, 5) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierInvoiceDate")
                                debitNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                                debitNoteArray(y, 7) = dsDebitNote.Tables(0).Rows(i).Item("dueDate2")
                                debitNoteArray(y, 8) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteSellingCurrencyCode")
                                debitNoteArray(y, 9) = totalCustomerDNBuying.ToString()
                                debitNoteArray(y, 10) = "0" 'VAT Number
                                debitNoteArray(y, 11) = "7" 'VAT Category
                                debitNoteArray(y, 12) = buyVat.ToString()
                                debitNoteArray(y, 13) = "2.1" 'INC_EXP_CATEGORY
                                debitNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                                debitNoteArray(y, 15) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteCreatedDate")
                                debitNoteArray(y, 16) = danaosUsername
                                debitNoteArray(y, 17) = "D" 'DEBIT CREDI
                                debitNoteArray(y, 18) = "" 'VESSEL_CODE
                                debitNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                                debitNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                                debitNoteArray(y, 21) = "40" 'JOURNALSERIES
                                debitNoteArray(y, 22) = "0" ' BOOK VALUE
                                debitNoteArray(y, 23) = "0" ' ENTRYRATE
                                debitNoteArray(y, 24) = "True"

                                ' insert total + vat
                                y += 1
                                'Inserts the First Row of the Amount without VAT
                                'invoicesBuyArray(u, 0) = dsBuyingInvoices.Tables("invoicesSelling").Rows(y).Item("deliveryID")
                                debitNoteArray(y, 0) = dsDebitNote.Tables(0).Rows(i).Item("deliveryMergedID")
                                debitNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                                debitNoteArray(y, 2) = dtData.Tables("outputvalues").Rows(0).Item("purchaseCode")
                                If dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierType_ID") = 1 Then
                                    debitNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("supplierCode")
                                ElseIf dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierType_ID") = 2 Then
                                    debitNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("forwardedCode")
                                ElseIf dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierType_ID") = 3 Then
                                    debitNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("clearingAgentCode")
                                End If
                                debitNoteArray(y, 4) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierInvoiceNumber")
                                debitNoteArray(y, 5) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierInvoiceDate")
                                debitNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                                debitNoteArray(y, 7) = dsDebitNote.Tables(0).Rows(i).Item("dueDate2")
                                debitNoteArray(y, 8) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteSellingCurrencyCode")
                                debitNoteArray(y, 9) = totalItemsBuying.ToString()
                                debitNoteArray(y, 10) = "0" 'VAT Number
                                debitNoteArray(y, 11) = "2" 'VAT Category
                                debitNoteArray(y, 12) = "0"
                                debitNoteArray(y, 13) = "" 'INC_EXP_CATEGORY
                                debitNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                                debitNoteArray(y, 15) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteCreatedDate")
                                debitNoteArray(y, 16) = danaosUsername
                                debitNoteArray(y, 17) = "C" 'DEBIT CREDI
                                debitNoteArray(y, 18) = "" 'VESSEL_CODE
                                debitNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                                debitNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                                debitNoteArray(y, 21) = "40" 'JOURNALSERIES
                                debitNoteArray(y, 22) = "0" ' BOOK VALUE
                                debitNoteArray(y, 23) = "0" ' ENTRYRATE
                                debitNoteArray(y, 24) = "True"

                                ' if vat is 0 skip dont insert anything
                                y += 1
                                'invoicesBuyArray(u, 0) = dsBuyingInvoices.Tables("invoicesSelling").Rows(y).Item("deliveryID")
                                debitNoteArray(y, 0) = dsDebitNote.Tables(0).Rows(i).Item("deliveryMergedID")
                                debitNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                                debitNoteArray(y, 2) = "38601"
                                debitNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("purchaseCode")
                                debitNoteArray(y, 4) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierInvoiceNumber")
                                debitNoteArray(y, 5) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierInvoiceDate")
                                debitNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                                debitNoteArray(y, 7) = dsDebitNote.Tables(0).Rows(i).Item("dueDate2")
                                debitNoteArray(y, 8) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteSellingCurrencyCode")
                                debitNoteArray(y, 9) = buyVat.ToString()
                                debitNoteArray(y, 10) = "0" 'VAT Number
                                debitNoteArray(y, 11) = "2" 'VAT Category
                                debitNoteArray(y, 12) = "0"
                                debitNoteArray(y, 13) = "" 'INC_EXP_CATEGORY
                                debitNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                                debitNoteArray(y, 15) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteCreatedDate")
                                debitNoteArray(y, 16) = danaosUsername
                                debitNoteArray(y, 17) = "C" 'DEBIT CREDI
                                debitNoteArray(y, 18) = "" 'VESSEL_CODE
                                debitNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                                debitNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                                debitNoteArray(y, 21) = "40" 'JOURNALSERIES
                                debitNoteArray(y, 22) = "0" ' BOOK VALUE                        
                                debitNoteArray(y, 23) = "0" ' ENTRYRATE
                                debitNoteArray(y, 24) = "True"

                            Else
                                If y > 0 Then y += 1

                                'invoicesBuyArray(u, 0) = dsBuyingInvoices.Tables("invoicesSelling").Rows(y).Item("deliveryID")
                                debitNoteArray(y, 0) = dsDebitNote.Tables(0).Rows(i).Item("deliveryMergedID")
                                debitNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                                If dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierType_ID") = 1 Then
                                    debitNoteArray(y, 2) = dtData.Tables("outputvalues").Rows(0).Item("supplierCode")
                                    debitNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("supplierCode")
                                ElseIf dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierType_ID") = 2 Then
                                    debitNoteArray(y, 2) = dsCustomerDN.Tables(0).Rows(i).Item("forwarderCode")
                                    debitNoteArray(y, 3) = dsCustomerDN.Tables(0).Rows(i).Item("forwarderCode")
                                ElseIf dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierType_ID") = 3 Then
                                    debitNoteArray(y, 2) = dsCustomerDN.Tables(0).Rows(i).Item("clearingAgentCode")
                                    debitNoteArray(y, 3) = dsCustomerDN.Tables(0).Rows(i).Item("clearingAgentCode")
                                End If
                                debitNoteArray(y, 4) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierInvoiceNumber")
                                debitNoteArray(y, 5) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierInvoiceDate")
                                debitNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                                debitNoteArray(y, 7) = dsDebitNote.Tables(0).Rows(i).Item("dueDate2")
                                debitNoteArray(y, 8) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteSellingCurrencyCode")
                                debitNoteArray(y, 9) = totalCustomerDNBuying.ToString()
                                debitNoteArray(y, 10) = "0" 'VAT Number
                                debitNoteArray(y, 11) = "2" 'VAT Category
                                debitNoteArray(y, 12) = buyVat.ToString()
                                debitNoteArray(y, 13) = "2.1" 'INC_EXP_CATEGORY
                                debitNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                                debitNoteArray(y, 15) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteCreatedDate")
                                debitNoteArray(y, 16) = danaosUsername
                                debitNoteArray(y, 17) = "D" 'DEBIT CREDI
                                debitNoteArray(y, 18) = "" 'VESSEL_CODE
                                debitNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                                debitNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                                debitNoteArray(y, 21) = "40" 'JOURNALSERIES
                                debitNoteArray(y, 22) = "0" ' BOOK VALUE
                                debitNoteArray(y, 23) = "0" ' ENTRYRATE
                                debitNoteArray(y, 24) = "True"

                                ' insert total + vat
                                y += 1
                                'Inserts the First Row of the Amount without VAT
                                'invoicesBuyArray(u, 0) = dsBuyingInvoices.Tables("invoicesSelling").Rows(y).Item("deliveryID")
                                debitNoteArray(y, 0) = dsDebitNote.Tables(0).Rows(i).Item("deliveryMergedID")
                                debitNoteArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                                debitNoteArray(y, 2) = dtData.Tables("outputvalues").Rows(0).Item("purchaseCode")
                                If dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierType_ID") = 1 Then
                                    debitNoteArray(y, 3) = dtData.Tables("outputvalues").Rows(0).Item("supplierCode")
                                ElseIf dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierType_ID") = 2 Then
                                    debitNoteArray(y, 3) = dsCustomerDN.Tables(0).Rows(i).Item("forwarderCode")
                                ElseIf dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierType_ID") = 3 Then
                                    debitNoteArray(y, 3) = dsCustomerDN.Tables(0).Rows(i).Item("clearingAgentCode")
                                End If
                                debitNoteArray(y, 4) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierInvoiceNumber")
                                debitNoteArray(y, 5) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteItemSupplierInvoiceDate")
                                debitNoteArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                                debitNoteArray(y, 7) = dsDebitNote.Tables(0).Rows(i).Item("dueDate2")
                                debitNoteArray(y, 8) = dsCustomerDN.Tables(0).Rows(i).Item("debitNoteSellingCurrencyCode")
                                debitNoteArray(y, 9) = totalCustomerDNBuying.ToString()
                                debitNoteArray(y, 10) = "0" 'VAT Number
                                debitNoteArray(y, 11) = "2" 'VAT Category
                                debitNoteArray(y, 12) = buyVat.ToString()
                                debitNoteArray(y, 13) = "" 'INC_EXP_CATEGORY
                                debitNoteArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                                debitNoteArray(y, 15) = dsDebitNote.Tables(0).Rows(i).Item("debitNoteCreatedDate")
                                debitNoteArray(y, 16) = danaosUsername
                                debitNoteArray(y, 17) = "C" 'DEBIT CREDI
                                debitNoteArray(y, 18) = "" 'VESSEL_CODE
                                debitNoteArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                                debitNoteArray(y, 20) = "023" 'JOURNAL_COMPANY
                                debitNoteArray(y, 21) = "40" 'JOURNALSERIES
                                debitNoteArray(y, 22) = "0" ' BOOK VALUE
                                debitNoteArray(y, 23) = "0" ' ENTRYRATE
                                debitNoteArray(y, 24) = "True"
                            End If
                        End If
                    End If
                End If
            Next
        End If


        Return debitNoteArray
        'DEBIT NOTE END
        '======================================================================================================================'
    End Function

    ''' <summary>
    ''' Function that exports rows of adjustments in Danaos
    ''' </summary>
    ''' <returns></returns>
    Public Function adjustmentNote() As String(,)
        Dim dsAdjustment = generaldataset.dsAdjustmentNote(orderID)
        Dim rowsCount = dsAdjustment.Tables(0).Rows.Count - 1
        Dim invRowsCount = (rowsCount + 1) * 6
        Dim adjustmentArray(invRowsCount, 24) As String
        Dim mergedID = 0
        Dim deliveryID = 0
        Dim amount = 0.00
        Dim currency = ""
        Dim vat = 0.00
        Dim suppInv = ""
        Dim supCode = ""
        Dim y As Integer = 0

        ' LOop into a dataset where there are adjustments created

        If dsAdjustment.Tables(0).Rows.Count > 0 Then
            For i = 0 To dsAdjustment.Tables(0).Rows.Count - 1
                If dsAdjustment.Tables(0).Rows(i).Item("exported") = 0 Then

                    Dim orclCmd As New OracleCommand

                    mergedID = dsAdjustment.Tables(0).Rows(i).Item("deliveryMergedID")
                    deliveryID = dsAdjustment.Tables(0).Rows(i).Item("deliveryID")
                    Dim tmpAdjID = dsAdjustment.Tables(0).Rows(i).Item("adjustmentID")

                    ' get the datya of the order
                    Dim dtData = generaldataset.loadOutputValues(orderID, mergedID)

                    If dtData.Tables("outputvalues").Rows(0).Item("vesselName") = "NA" Then
                        dtData.Tables("outputvalues").Rows(0).Item("vesselName") = "NOT SPECIFIED"
                    End If

                    ' set the narrative so we can get the data from danaos
                    Dim narrative = dtData.Tables("outputvalues").Rows(0).Item("narativeOfEntry")

                    Try
                        orclConn.Close()
                    Catch ex As Exception

                    End Try


                    Try
                        orclConn.Open()
                    Catch ex As Exception

                    End Try

                    Dim adjustmentNarr = ""

                    ' get the data from danaos into a datatable
                    orclCmd.Connection = orclConn

                    'orclCmd.CommandText = "SELECT * FROM DOCUMENTS WHERE NARRATIVE like '%" & narrative & "%'"
                    Try
                        If dsAdjustment.Tables(0).Rows(i).Item("forwarder_ID") <> 0 Then
                            For z = 0 To dtData.Tables(0).Rows.Count - 1
                                If dtData.Tables("outputvalues").Rows(z).Item("deliveryID") = dsAdjustment.Tables(0).Rows(i).Item("deliveryID") Then
                                    adjustmentNarr = TruncString.Truncate(dtData.Tables("outputvalues").Rows(z).Item("forwarderCompName"), 25)
                                End If
                            Next
                            orclCmd.CommandText = "SELECT * FROM DOCUMENTS WHERE NARRATIVE like '%" & narrative & "%" & adjustmentNarr & "%' AND (SERIAL_NO like '%??INV-%' OR Serial_NO like '%????%') ORDER BY SERIAL_NO ASC, DEBIT_CREDIT ASC"
                        ElseIf dsAdjustment.Tables(0).Rows(i).Item("clearingAgent_ID") <> 0 Then
                            For z = 0 To dtData.Tables(0).Rows.Count - 1
                                If dtData.Tables("outputvalues").Rows(z).Item("deliveryID") = dsAdjustment.Tables(0).Rows(i).Item("deliveryID") Then
                                    adjustmentNarr = TruncString.Truncate(dtData.Tables("outputvalues").Rows(z).Item("clearingAgentCompName"), 25)
                                End If
                            Next
                            orclCmd.CommandText = "SELECT * FROM DOCUMENTS WHERE NARRATIVE like '%" & narrative & "%" & adjustmentNarr & "%' AND (SERIAL_NO like '%??INV-%' OR Serial_NO like '%????%') ORDER BY SERIAL_NO ASC, DEBIT_CREDIT ASC"
                        ElseIf dsAdjustment.Tables(0).Rows(i).Item("generalSupplier1_ID") <> 0 Then
                            For z = 0 To dtData.Tables(0).Rows.Count - 1
                                If dtData.Tables("outputvalues").Rows(z).Item("deliveryID") = dsAdjustment.Tables(0).Rows(i).Item("deliveryID") Then
                                    adjustmentNarr = TruncString.Truncate(dtData.Tables("outputvalues").Rows(z).Item("generalSupplier1CompName"), 25)
                                End If
                            Next
                            orclCmd.CommandText = "SELECT * FROM DOCUMENTS WHERE NARRATIVE like '%" & narrative & "%" & adjustmentNarr & "%' AND (SERIAL_NO like '%??INV-%' OR Serial_NO like '%????%') ORDER BY SERIAL_NO ASC, DEBIT_CREDIT ASC"
                        ElseIf dsAdjustment.Tables(0).Rows(i).Item("generalSupplier2_ID") <> 0 Then
                            For z = 0 To dtData.Tables(0).Rows.Count - 1
                                If dtData.Tables("outputvalues").Rows(z).Item("deliveryID") = dsAdjustment.Tables(0).Rows(i).Item("deliveryID") Then
                                    adjustmentNarr = TruncString.Truncate(dtData.Tables("outputvalues").Rows(z).Item("generalSupplier2CompName"), 25)
                                End If
                            Next
                            orclCmd.CommandText = "SELECT * FROM DOCUMENTS WHERE NARRATIVE like '%" & narrative & "%" & adjustmentNarr & "%' AND (SERIAL_NO like '%??INV-%' OR Serial_NO like '%????%') ORDER BY SERIAL_NO ASC, DEBIT_CREDIT ASC"
                        ElseIf dsAdjustment.Tables(0).Rows(i).Item("generalSupplier3_ID") <> 0 Then
                            For z = 0 To dtData.Tables(0).Rows.Count - 1
                                If dtData.Tables("outputvalues").Rows(z).Item("deliveryID") = dsAdjustment.Tables(0).Rows(i).Item("deliveryID") Then
                                    adjustmentNarr = TruncString.Truncate(dtData.Tables("outputvalues").Rows(z).Item("generalSupplier3CompName"), 25)
                                End If
                            Next
                            orclCmd.CommandText = "SELECT * FROM DOCUMENTS WHERE NARRATIVE like '%" & narrative & "%" & adjustmentNarr & "%' AND (SERIAL_NO like '%??INV-%' OR Serial_NO like '%????%') ORDER BY SERIAL_NO ASC, DEBIT_CREDIT ASC"
                        End If
                    Catch ex As Exception
                        orclCmd.CommandText = "SELECT * FROM DOCUMENTS WHERE NARRATIVE like '%" & narrative & "%' AND (SERIAL_NO like '%??INV-%' OR Serial_NO like '%????%') ORDER BY SERIAL_NO ASC, DEBIT_CREDIT, FROM_TO ASC"
                    End Try

                    'orclCmd.CommandText = "SELECT * FROM DOCUMENTS WHERE NARRATIVE like '%" & narrative & "%' AND (SERIAL_NO like '%??INV-%' OR Serial_NO like '%????%') ORDER BY SERIAL_NO ASC, DEBIT_CREDIT, FROM_TO ASC"

                    orclCmd.CommandType = CommandType.Text
                    ' orclConn.Open()
                    Dim dr As New DataTable
                    ' Dim orclName As String
                    dr.Load(orclCmd.ExecuteReader())

                    If y < 5 And y <> 0 Then y += 1

                    ' fill the table with the data
                    For o = 0 To dr.Rows.Count - 1
                        Select Case o
                            Case 0
                                adjustmentArray(y, 0) = mergedID
                                adjustmentArray(y, 1) = dr.Rows(o).Item("OUR_COMPANY").ToString()
                                adjustmentArray(y, 2) = dr.Rows(o).Item("FROM_TO").ToString()
                                adjustmentArray(y, 3) = "39002"
                                adjustmentArray(y, 4) = dr.Rows(o).Item("SERIAL_NO").ToString()
                                adjustmentArray(y, 5) = dr.Rows(o).Item("ISSUE_DATE").ToString()
                                adjustmentArray(y, 6) = dr.Rows(o).Item("NARRATIVE").ToString()
                                adjustmentArray(y, 7) = dr.Rows(o).Item("DUE_DATE").ToString()
                                adjustmentArray(y, 8) = dr.Rows(o).Item("CURRENCY").ToString()
                                adjustmentArray(y, 9) = dr.Rows(o).Item("AMOUNT").ToString()
                                adjustmentArray(y, 10) = dr.Rows(o).Item("VAT_NUMBER").ToString()
                                adjustmentArray(y, 11) = dr.Rows(o).Item("VAT_CATEGORY").ToString()
                                adjustmentArray(y, 12) = dr.Rows(o).Item("VAT_AMOUNT").ToString()
                                adjustmentArray(y, 13) = dr.Rows(o).Item("INC_EXP_CATEGORY").ToString()
                                adjustmentArray(y, 14) = dr.Rows(o).Item("PLACE").ToString()
                                adjustmentArray(y, 15) = dsAdjustment.Tables(0).Rows(i).Item("adjustmentDate")
                                adjustmentArray(y, 16) = dr.Rows(o).Item("JOURNAL_USER").ToString()
                                If dr.Rows(o).Item("DEBIT_CREDIT").ToString() = "C" Then
                                    adjustmentArray(y, 17) = "D"
                                ElseIf dr.Rows(o).Item("DEBIT_CREDIT").ToString() = "D" Then
                                    adjustmentArray(y, 17) = "C"
                                End If
                                adjustmentArray(y, 18) = dr.Rows(o).Item("VESSEL_CODE").ToString()
                                adjustmentArray(y, 19) = narrative
                                adjustmentArray(y, 20) = dr.Rows(o).Item("JOURNAL_COMPANY").ToString()
                                adjustmentArray(y, 21) = "40"
                                adjustmentArray(y, 22) = dr.Rows(o).Item("BOOK_VALUE").ToString()
                                adjustmentArray(y, 23) = dr.Rows(o).Item("ENTRY_RATE").ToString()
                                adjustmentArray(y, 24) = "Reversed"
                                y = y + 1

                                If (rowsCount + 1 = 2 And o = 1) Then
                                    GoTo endReverse
                                End If
                            Case 1
                                If dr.Rows(0).Item("DEBIT_CREDIT").ToString() <> dr.Rows(1).Item("DEBIT_CREDIT").ToString() Then
                                    adjustmentArray(y, 0) = mergedID
                                    adjustmentArray(y, 1) = dr.Rows(o).Item("OUR_COMPANY").ToString()
                                    adjustmentArray(y, 2) = dr.Rows(o).Item("FROM_TO").ToString()
                                    adjustmentArray(y, 3) = "39002"
                                    adjustmentArray(y, 4) = dr.Rows(o).Item("SERIAL_NO").ToString()
                                    adjustmentArray(y, 5) = dr.Rows(o).Item("ISSUE_DATE").ToString()
                                    adjustmentArray(y, 6) = dr.Rows(o).Item("NARRATIVE").ToString()
                                    adjustmentArray(y, 7) = dr.Rows(o).Item("DUE_DATE").ToString()
                                    adjustmentArray(y, 8) = dr.Rows(o).Item("CURRENCY").ToString()
                                    adjustmentArray(y, 9) = dr.Rows(o).Item("AMOUNT").ToString()
                                    adjustmentArray(y, 10) = dr.Rows(o).Item("VAT_NUMBER").ToString()
                                    adjustmentArray(y, 11) = dr.Rows(o).Item("VAT_CATEGORY").ToString()
                                    adjustmentArray(y, 12) = dr.Rows(o).Item("VAT_AMOUNT").ToString()
                                    adjustmentArray(y, 13) = dr.Rows(o).Item("INC_EXP_CATEGORY").ToString()
                                    adjustmentArray(y, 14) = dr.Rows(o).Item("PLACE").ToString()
                                    adjustmentArray(y, 15) = dsAdjustment.Tables(0).Rows(i).Item("adjustmentDate")
                                    adjustmentArray(y, 16) = dr.Rows(o).Item("JOURNAL_USER").ToString()
                                    If dr.Rows(o).Item("DEBIT_CREDIT").ToString() = "C" Then
                                        adjustmentArray(y, 17) = "D"
                                    ElseIf dr.Rows(o).Item("DEBIT_CREDIT").ToString() = "D" Then
                                        adjustmentArray(y, 17) = "C"
                                    End If
                                    adjustmentArray(y, 18) = dr.Rows(o).Item("VESSEL_CODE").ToString()
                                    adjustmentArray(y, 19) = narrative
                                    adjustmentArray(y, 20) = dr.Rows(o).Item("JOURNAL_COMPANY").ToString()
                                    adjustmentArray(y, 21) = "40"
                                    adjustmentArray(y, 22) = dr.Rows(o).Item("BOOK_VALUE").ToString()
                                    adjustmentArray(y, 23) = dr.Rows(o).Item("ENTRY_RATE").ToString()
                                    adjustmentArray(y, 24) = "Reversed"
                                    y = y + 1

                                    If (rowsCount + 1 = 2 And o = 1) Then
                                        GoTo endReverse
                                    End If
                                End If
                            Case 2
                                adjustmentArray(y, 0) = mergedID
                                adjustmentArray(y, 1) = dr.Rows(o).Item("OUR_COMPANY").ToString()
                                adjustmentArray(y, 2) = dr.Rows(o).Item("FROM_TO").ToString()
                                adjustmentArray(y, 3) = "39002"
                                adjustmentArray(y, 4) = dr.Rows(o).Item("SERIAL_NO").ToString()
                                adjustmentArray(y, 5) = dr.Rows(o).Item("ISSUE_DATE").ToString()
                                adjustmentArray(y, 6) = dr.Rows(o).Item("NARRATIVE").ToString()
                                adjustmentArray(y, 7) = dr.Rows(o).Item("DUE_DATE").ToString()
                                adjustmentArray(y, 8) = dr.Rows(o).Item("CURRENCY").ToString()
                                adjustmentArray(y, 9) = dr.Rows(o).Item("AMOUNT").ToString()
                                adjustmentArray(y, 10) = dr.Rows(o).Item("VAT_NUMBER").ToString()
                                adjustmentArray(y, 11) = dr.Rows(o).Item("VAT_CATEGORY").ToString()
                                adjustmentArray(y, 12) = dr.Rows(o).Item("VAT_AMOUNT").ToString()
                                adjustmentArray(y, 13) = dr.Rows(o).Item("INC_EXP_CATEGORY").ToString()
                                adjustmentArray(y, 14) = dr.Rows(o).Item("PLACE").ToString()
                                adjustmentArray(y, 15) = dsAdjustment.Tables(0).Rows(i).Item("adjustmentDate")
                                adjustmentArray(y, 16) = dr.Rows(o).Item("JOURNAL_USER").ToString()
                                If dr.Rows(o).Item("DEBIT_CREDIT").ToString() = "C" Then
                                    adjustmentArray(y, 17) = "D"
                                ElseIf dr.Rows(o).Item("DEBIT_CREDIT").ToString() = "D" Then
                                    adjustmentArray(y, 17) = "C"
                                End If
                                adjustmentArray(y, 18) = dr.Rows(o).Item("VESSEL_CODE").ToString()
                                adjustmentArray(y, 19) = narrative
                                adjustmentArray(y, 20) = dr.Rows(o).Item("JOURNAL_COMPANY").ToString()
                                adjustmentArray(y, 21) = "40"
                                adjustmentArray(y, 22) = dr.Rows(o).Item("BOOK_VALUE").ToString()
                                adjustmentArray(y, 23) = dr.Rows(o).Item("ENTRY_RATE").ToString()
                                adjustmentArray(y, 24) = "Reversed"
                                y = y + 1

                                If (rowsCount + 1 = 2 And o = 1) Then
                                    GoTo endReverse
                                End If
                        End Select
                    Next

endReverse:

                    Try
                        orclConn.Close()
                    Catch ex As Exception

                    End Try

                    ' if its a forwarder adjustment
                    If dsAdjustment.Tables(0).Rows(i).Item("forwarder_ID") <> 0 Then
                        ' get the amount, vat, currency, supp inv, supcode for the export
                        SQL = "SELECT freightBuying,
													(SELECT vatValue from easyenquiry.easyenquiry.vat where vatID = freightBuyingVat_ID) as freightVAT,
													freightBuyingCurr,
                                                    forwarderInvoiceNumber,
													(Select forwarderCode From easyenquiry.easyenquiry.forwarders where forwarder_ID = forwarderID) as forwarderCode

													From easyenquiry.easyenquiry.delivery

													LEFT JOIN easyenquiry.easyenquiry.suppliersinvoices On deliveryID = delivery_ID

													Where deliveryID = '" & deliveryID & "'"
                        myCommand.Connection = conn
                        conn.Close()
                        conn.Open()
                        myCommand.CommandText = SQL
                        myReader = myCommand.ExecuteReader
                        If (myReader.HasRows = True) Then
                            While myReader.Read
                                amount = myReader.GetValue(myReader.GetOrdinal("freightBuying"))
                                vat = myReader.GetValue(myReader.GetOrdinal("freightVAT"))
                                currency = myReader.GetValue(myReader.GetOrdinal("freightBuyingCurr"))
                                suppInv = myReader.GetValue(myReader.GetOrdinal("forwarderInvoiceNumber"))
                                supCode = myReader.GetValue(myReader.GetOrdinal("forwarderCode"))
                            End While
                            myReader.Close()
                        End If
                        conn.Close()
                        ' if its a clearing agent adjustment
                    ElseIf dsAdjustment.Tables(0).Rows(i).Item("clearingAgent_ID") <> 0 Then
                        ' get the amount, vat, currency, supp inv, supcode for the export
                        SQL = "Select customsBuying,
													(Select vatValue from easyenquiry.easyenquiry.vat where vatID = customsBuyingVat_ID) as customsVAT,
													customsBuyingCurr,
													clearingAgentInvoiceNumber,
													(Select clearingAgentCode From easyenquiry.easyenquiry.clearingagents where clearingAgent_ID = clearingAgentID) as clearingAgentCode

													From easyenquiry.easyenquiry.delivery
													LEFT JOIN easyenquiry.easyenquiry.suppliersinvoices On deliveryID = delivery_ID
													Where deliveryID = '" & deliveryID & "'"
                        myCommand.Connection = conn
                        conn.Close()
                        conn.Open()
                        myCommand.CommandText = SQL
                        myReader = myCommand.ExecuteReader
                        If (myReader.HasRows = True) Then
                            While myReader.Read
                                amount = myReader.GetValue(myReader.GetOrdinal("customsBuying"))
                                vat = myReader.GetValue(myReader.GetOrdinal("customsVAT"))
                                currency = myReader.GetValue(myReader.GetOrdinal("customsBuyingCurr"))
                                suppInv = myReader.GetValue(myReader.GetOrdinal("clearingAgentInvoiceNumber"))
                                supCode = myReader.GetValue(myReader.GetOrdinal("clearingAgentCode"))
                            End While
                            myReader.Close()
                        End If
                        conn.Close()

                        ' if its a general supplier adjustment 
                    ElseIf dsAdjustment.Tables(0).Rows(i).Item("generalSupplier1_ID") <> 0 Then

                        SQL = "Select extraCharges1Buying,
													(Select vatValue from easyenquiry.easyenquiry.vat where vatID = extraCharges1BuyingVat_ID) as ec1VAT,
													extraCharges1BuyingCurr,
													generalSupplier1InvoiceNumber,
													(Select generalSupplierCode From easyenquiry.easyenquiry.generalsuppliers where generalSupplier1_ID = generalSupplierID) as genSup1Code

													From easyenquiry.easyenquiry.delivery
													LEFT JOIN easyenquiry.easyenquiry.suppliersinvoices On deliveryID = delivery_ID

													Where deliveryID = '" & deliveryID & "'"
                        myCommand.Connection = conn
                        conn.Close()
                        conn.Open()
                        myCommand.CommandText = SQL
                        myReader = myCommand.ExecuteReader
                        If (myReader.HasRows = True) Then
                            While myReader.Read
                                amount = myReader.GetValue(myReader.GetOrdinal("extraCharges1Buying"))
                                vat = myReader.GetValue(myReader.GetOrdinal("ec1VAT"))
                                currency = myReader.GetValue(myReader.GetOrdinal("extraCharges1BuyingCurr"))
                                suppInv = myReader.GetValue(myReader.GetOrdinal("generalSupplier1InvoiceNumber"))
                                supCode = myReader.GetValue(myReader.GetOrdinal("genSup1Code"))
                            End While
                            myReader.Close()
                        End If
                        conn.Close()

                    ElseIf dsAdjustment.Tables(0).Rows(i).Item("generalSupplier2_ID") <> 0 Then

                        SQL = "Select extraCharges2Buying,
													(Select vatValue from easyenquiry.easyenquiry.vat where vatID = extraCharges2BuyingVat_ID) as ec2VAT,
													extraCharges2BuyingCurr,
													generalSupplier2InvoiceNumber,
													(Select generalSupplierCode From easyenquiry.easyenquiry.generalsuppliers where generalSupplier2_ID = generalSupplierID) as genSup2Code

													From easyenquiry.easyenquiry.delivery
													LEFT JOIN easyenquiry.easyenquiry.suppliersinvoices On deliveryID = delivery_ID

													Where deliveryID = '" & deliveryID & "'"
                        myCommand.Connection = conn
                        conn.Close()
                        conn.Open()
                        myCommand.CommandText = SQL
                        myReader = myCommand.ExecuteReader
                        If (myReader.HasRows = True) Then
                            While myReader.Read
                                amount = myReader.GetValue(myReader.GetOrdinal("extraCharges2Buying"))
                                vat = myReader.GetValue(myReader.GetOrdinal("ec2VAT"))
                                currency = myReader.GetValue(myReader.GetOrdinal("extraCharges2BuyingCurr"))
                                suppInv = myReader.GetValue(myReader.GetOrdinal("generalSupplier2InvoiceNumber"))
                                supCode = myReader.GetValue(myReader.GetOrdinal("genSup2Code"))

                            End While
                            myReader.Close()
                        End If
                        conn.Close()
                    ElseIf dsAdjustment.Tables(0).Rows(i).Item("generalSupplier3_ID") <> 0 Then
                        SQL = "Select extraCharges3Buying,
													(Select vatValue from easyenquiry.easyenquiry.vat where vatID = extraCharges3BuyingVat_ID) as ec3VAT,
													extraCharges3BuyingCurr,
													generalSupplier3InvoiceNumber,
													(Select generalSupplierCode From easyenquiry.easyenquiry.generalsuppliers where generalSupplier3_ID = generalSupplierID) as genSup3Code

													From easyenquiry.easyenquiry.delivery
													LEFT JOIN easyenquiry.easyenquiry.suppliersinvoices On deliveryID = delivery_ID

													Where deliveryID = '" & deliveryID & "'"
                        myCommand.Connection = conn
                        conn.Close()
                        conn.Open()
                        myCommand.CommandText = SQL
                        myReader = myCommand.ExecuteReader
                        If (myReader.HasRows = True) Then
                            While myReader.Read
                                amount = myReader.GetValue(myReader.GetOrdinal("extraCharges3Buying"))
                                vat = myReader.GetValue(myReader.GetOrdinal("ec3VAT"))
                                currency = myReader.GetValue(myReader.GetOrdinal("extraCharges3BuyingCurr"))
                                suppInv = myReader.GetValue(myReader.GetOrdinal("generalSupplier3InvoiceNumber"))
                                supCode = myReader.GetValue(myReader.GetOrdinal("genSup3Code"))
                            End While
                            myReader.Close()
                        End If
                        conn.Close()
                    End If

                    narrative = dtData.Tables("outputvalues").Rows(0).Item("narativeOfEntry") + "-" + dtData.Tables("outputvalues").Rows(0).Item("vesselName") + "-ADJUSTMENT NOTE"

                    If dsAdjustment.Tables(0).Rows(i).Item("adjustmentVatLS") > 0.00 Then
                        If y > 0 Then y += 1

                        'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                        adjustmentArray(y, 0) = dsAdjustment.Tables(0).Rows(i).Item("deliveryMergedID")
                        adjustmentArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                        adjustmentArray(y, 2) = supCode
                        adjustmentArray(y, 3) = supCode
                        adjustmentArray(y, 4) = suppInv
                        adjustmentArray(y, 5) = dsAdjustment.Tables(0).Rows(i).Item("adjustmentDate")
                        adjustmentArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                        If dsAdjustment.Tables(0).Rows(i).Item("forwarder_ID") <> 0 Then
                            adjustmentArray(y, 7) = dsAdjustment.Tables(0).Rows(i).Item("dueDateF")
                        ElseIf dsAdjustment.Tables(0).Rows(i).Item("generalSupplier1_ID") <> 0 Then
                            adjustmentArray(y, 7) = dsAdjustment.Tables(0).Rows(i).Item("dueDateGS1")
                        ElseIf dsAdjustment.Tables(0).Rows(i).Item("generalSupplier2_ID") <> 0 Then
                            adjustmentArray(y, 7) = dsAdjustment.Tables(0).Rows(i).Item("dueDateGS2")
                        ElseIf dsAdjustment.Tables(0).Rows(i).Item("generalSupplier3_ID") <> 0 Then
                            adjustmentArray(y, 7) = dsAdjustment.Tables(0).Rows(i).Item("dueDateGS3")
                        Else
                            adjustmentArray(y, 7) = dsAdjustment.Tables(0).Rows(i).Item("dueDate2")
                        End If
                        adjustmentArray(y, 8) = currency
                        adjustmentArray(y, 9) = amount + dsAdjustment.Tables(0).Rows(i).Item("adjustmentVatLS")
                        adjustmentArray(y, 10) = "0" 'VAT Number
                        adjustmentArray(y, 11) = "2" 'VAT Category
                        adjustmentArray(y, 12) = "0"
                        adjustmentArray(y, 13) = "2.1" 'INC_EXP_CATEGORY
                        adjustmentArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                        adjustmentArray(y, 15) = dsAdjustment.Tables(0).Rows(i).Item("adjustmentDate")
                        adjustmentArray(y, 16) = danaosUsername
                        adjustmentArray(y, 17) = "C" 'DEBIT CREDIT
                        adjustmentArray(y, 18) = "" 'VESSEL_CODE
                        adjustmentArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                        adjustmentArray(y, 20) = "023" 'JOURNAL_COMPANY
                        adjustmentArray(y, 21) = "40" 'JOURNALSERIES
                        adjustmentArray(y, 22) = "0" ' BOOK VALUE
                        adjustmentArray(y, 23) = "0" ' ENTRYRATE
                        adjustmentArray(y, 24) = "True" ' ENTRYRATE
                        y += 1

                        'Inserts the First Row of the Amount without VAT
                        'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                        adjustmentArray(y, 0) = dsAdjustment.Tables(0).Rows(i).Item("deliveryMergedID")
                        adjustmentArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                        adjustmentArray(y, 2) = dtData.Tables("outputvalues").Rows(0).Item("purchaseCode")
                        adjustmentArray(y, 3) = supCode
                        adjustmentArray(y, 4) = suppInv
                        adjustmentArray(y, 5) = dsAdjustment.Tables(0).Rows(i).Item("adjustmentDate")
                        adjustmentArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                        If dsAdjustment.Tables(0).Rows(i).Item("forwarder_ID") <> 0 Then
                            adjustmentArray(y, 7) = dsAdjustment.Tables(0).Rows(i).Item("dueDateF")
                        ElseIf dsAdjustment.Tables(0).Rows(i).Item("generalSupplier1_ID") <> 0 Then
                            adjustmentArray(y, 7) = dsAdjustment.Tables(0).Rows(i).Item("dueDateGS1")
                        ElseIf dsAdjustment.Tables(0).Rows(i).Item("generalSupplier2_ID") <> 0 Then
                            adjustmentArray(y, 7) = dsAdjustment.Tables(0).Rows(i).Item("dueDateGS2")
                        ElseIf dsAdjustment.Tables(0).Rows(i).Item("generalSupplier3_ID") <> 0 Then
                            adjustmentArray(y, 7) = dsAdjustment.Tables(0).Rows(i).Item("dueDateGS3")
                        Else
                            adjustmentArray(y, 7) = dsAdjustment.Tables(0).Rows(i).Item("dueDate2")
                        End If
                        adjustmentArray(y, 8) = currency
                        adjustmentArray(y, 9) = amount.ToString()
                        adjustmentArray(y, 10) = "0" 'VAT Number
                        adjustmentArray(y, 11) = "2" 'VAT Category
                        adjustmentArray(y, 12) = "0"
                        adjustmentArray(y, 13) = "" 'INC_EXP_CATEGORY
                        adjustmentArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                        adjustmentArray(y, 15) = dsAdjustment.Tables(0).Rows(i).Item("adjustmentDate")
                        adjustmentArray(y, 16) = danaosUsername
                        adjustmentArray(y, 17) = "D" 'DEBIT CREDI
                        adjustmentArray(y, 18) = "" 'VESSEL_CODE
                        adjustmentArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                        adjustmentArray(y, 20) = "023" 'JOURNAL_COMPANY
                        adjustmentArray(y, 21) = "40" 'JOURNALSERIES
                        adjustmentArray(y, 22) = "0" ' BOOK VALUE
                        adjustmentArray(y, 23) = "0" ' ENTRYRATE
                        adjustmentArray(y, 24) = "True" ' ENTRYRATE
                        y += 1

                        'Inserts the First Row of the Amount without VAT
                        'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                        adjustmentArray(y, 0) = dsAdjustment.Tables(0).Rows(i).Item("deliveryMergedID")
                        adjustmentArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                        adjustmentArray(y, 2) = "38601"
                        adjustmentArray(y, 3) = supCode
                        adjustmentArray(y, 4) = suppInv
                        adjustmentArray(y, 5) = dsAdjustment.Tables(0).Rows(i).Item("adjustmentDate")
                        adjustmentArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                        If dsAdjustment.Tables(0).Rows(i).Item("forwarder_ID") <> 0 Then
                            adjustmentArray(y, 7) = dsAdjustment.Tables(0).Rows(i).Item("dueDateF")
                        ElseIf dsAdjustment.Tables(0).Rows(i).Item("generalSupplier1_ID") <> 0 Then
                            adjustmentArray(y, 7) = dsAdjustment.Tables(0).Rows(i).Item("dueDateGS1")
                        ElseIf dsAdjustment.Tables(0).Rows(i).Item("generalSupplier2_ID") <> 0 Then
                            adjustmentArray(y, 7) = dsAdjustment.Tables(0).Rows(i).Item("dueDateGS2")
                        ElseIf dsAdjustment.Tables(0).Rows(i).Item("generalSupplier3_ID") <> 0 Then
                            adjustmentArray(y, 7) = dsAdjustment.Tables(0).Rows(i).Item("dueDateGS3")
                        Else
                            adjustmentArray(y, 7) = dsAdjustment.Tables(0).Rows(i).Item("dueDate2")
                        End If
                        adjustmentArray(y, 8) = currency
                        adjustmentArray(y, 9) = dsAdjustment.Tables(0).Rows(i).Item("adjustmentVatLS")
                        adjustmentArray(y, 10) = "0" 'VAT Number
                        adjustmentArray(y, 11) = "2" 'VAT Category
                        adjustmentArray(y, 12) = "0"
                        adjustmentArray(y, 13) = "" 'INC_EXP_CATEGORY
                        adjustmentArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                        adjustmentArray(y, 15) = dsAdjustment.Tables(0).Rows(i).Item("adjustmentDate")
                        adjustmentArray(y, 16) = danaosUsername
                        adjustmentArray(y, 17) = "D" 'DEBIT CREDI
                        adjustmentArray(y, 18) = "" 'VESSEL_CODE
                        adjustmentArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                        adjustmentArray(y, 20) = "023" 'JOURNAL_COMPANY
                        adjustmentArray(y, 21) = "40" 'JOURNALSERIES
                        adjustmentArray(y, 22) = "0" ' BOOK VALUE
                        adjustmentArray(y, 23) = "0" ' ENTRYRATE
                        adjustmentArray(y, 24) = "True" ' ENTRYRATE
                    Else
                        If y > 0 Then y += 1

                        'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                        adjustmentArray(y, 0) = dsAdjustment.Tables(0).Rows(i).Item("deliveryMergedID")
                        adjustmentArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                        adjustmentArray(y, 2) = supCode
                        adjustmentArray(y, 3) = supCode
                        adjustmentArray(y, 4) = suppInv
                        adjustmentArray(y, 5) = dsAdjustment.Tables(0).Rows(i).Item("adjustmentDate")
                        adjustmentArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                        If dsAdjustment.Tables(0).Rows(i).Item("forwarder_ID") <> 0 Then
                            adjustmentArray(y, 7) = dsAdjustment.Tables(0).Rows(i).Item("dueDateF")
                        ElseIf dsAdjustment.Tables(0).Rows(i).Item("generalSupplier1_ID") <> 0 Then
                            adjustmentArray(y, 7) = dsAdjustment.Tables(0).Rows(i).Item("dueDateGS1")
                        ElseIf dsAdjustment.Tables(0).Rows(i).Item("generalSupplier2_ID") <> 0 Then
                            adjustmentArray(y, 7) = dsAdjustment.Tables(0).Rows(i).Item("dueDateGS2")
                        ElseIf dsAdjustment.Tables(0).Rows(i).Item("generalSupplier3_ID") <> 0 Then
                            adjustmentArray(y, 7) = dsAdjustment.Tables(0).Rows(i).Item("dueDateGS3")
                        Else
                            adjustmentArray(y, 7) = dsAdjustment.Tables(0).Rows(i).Item("dueDate2")
                        End If
                        adjustmentArray(y, 8) = currency
                        adjustmentArray(y, 9) = amount.ToString()
                        adjustmentArray(y, 10) = "0" 'VAT Number
                        adjustmentArray(y, 11) = "2" 'VAT Category
                        adjustmentArray(y, 12) = "0"
                        adjustmentArray(y, 13) = "2.1" 'INC_EXP_CATEGORY
                        adjustmentArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                        adjustmentArray(y, 15) = dsAdjustment.Tables(0).Rows(i).Item("adjustmentDate")
                        adjustmentArray(y, 16) = danaosUsername
                        adjustmentArray(y, 17) = "C" 'DEBIT CREDIT
                        adjustmentArray(y, 18) = "" 'VESSEL_CODE
                        adjustmentArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                        adjustmentArray(y, 20) = "023" 'JOURNAL_COMPANY
                        adjustmentArray(y, 21) = "40" 'JOURNALSERIES
                        adjustmentArray(y, 22) = "0" ' BOOK VALUE
                        adjustmentArray(y, 23) = "0" ' ENTRYRATE
                        adjustmentArray(y, 24) = "True" ' ENTRYRATE
                        y += 1

                        'Inserts the First Row of the Amount without VAT
                        'invoicesSellArray(j, 0) = dsSellingInvoices.Tables("invoicesSelling").Rows(i).Item("delivery_ID")
                        adjustmentArray(y, 0) = dsAdjustment.Tables(0).Rows(i).Item("deliveryMergedID")
                        adjustmentArray(y, 1) = dtData.Tables("outputvalues").Rows(0).Item("companyCode")
                        adjustmentArray(y, 2) = dtData.Tables("outputvalues").Rows(0).Item("purchaseCode")
                        adjustmentArray(y, 3) = supCode
                        adjustmentArray(y, 4) = suppInv
                        adjustmentArray(y, 5) = dsAdjustment.Tables(0).Rows(i).Item("adjustmentDate")
                        adjustmentArray(y, 6) = TruncString.Truncate(narrative.ToString(), 60)
                        If dsAdjustment.Tables(0).Rows(i).Item("forwarder_ID") <> 0 Then
                            adjustmentArray(y, 7) = dsAdjustment.Tables(0).Rows(i).Item("dueDateF")
                        ElseIf dsAdjustment.Tables(0).Rows(i).Item("generalSupplier1_ID") <> 0 Then
                            adjustmentArray(y, 7) = dsAdjustment.Tables(0).Rows(i).Item("dueDateGS1")
                        ElseIf dsAdjustment.Tables(0).Rows(i).Item("generalSupplier2_ID") <> 0 Then
                            adjustmentArray(y, 7) = dsAdjustment.Tables(0).Rows(i).Item("dueDateGS2")
                        ElseIf dsAdjustment.Tables(0).Rows(i).Item("generalSupplier3_ID") <> 0 Then
                            adjustmentArray(y, 7) = dsAdjustment.Tables(0).Rows(i).Item("dueDateGS3")
                        Else
                            adjustmentArray(y, 7) = dsAdjustment.Tables(0).Rows(i).Item("dueDate2")
                        End If
                        adjustmentArray(y, 8) = currency
                        adjustmentArray(y, 9) = amount.ToString()
                        adjustmentArray(y, 10) = "0" 'VAT Number
                        adjustmentArray(y, 11) = "2" 'VAT Category
                        adjustmentArray(y, 12) = "0"
                        adjustmentArray(y, 13) = "" 'INC_EXP_CATEGORY
                        adjustmentArray(y, 14) = dtData.Tables("outputvalues").Rows(0).Item("deliveryPlace")
                        adjustmentArray(y, 15) = dsAdjustment.Tables(0).Rows(i).Item("adjustmentDate")
                        adjustmentArray(y, 16) = danaosUsername
                        adjustmentArray(y, 17) = "D" 'DEBIT CREDI
                        adjustmentArray(y, 18) = "" 'VESSEL_CODE
                        adjustmentArray(y, 19) = dtData.Tables("outputvalues").Rows(0).Item("originalRefNum")
                        adjustmentArray(y, 20) = "023" 'JOURNAL_COMPANY
                        adjustmentArray(y, 21) = "40" 'JOURNALSERIES
                        adjustmentArray(y, 22) = "0" ' BOOK VALUE
                        adjustmentArray(y, 23) = "0" ' ENTRYRATE
                        adjustmentArray(y, 24) = "True" ' ENTRYRATE
                    End If

                    anExported(tmpAdjID)

                End If
            Next
        End If

        Return adjustmentArray
    End Function

    ''' <summary>
    ''' check if finalized
    ''' </summary>
    ''' <param name="finalized"></param>
    ''' <returns></returns>
    Private Function finalizedBool(ByVal finalized As DataSet) As Boolean
        If finalized.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' check if CN exported
    ''' </summary>
    ''' <param name="creditnote"></param>
    ''' <returns></returns>
    Private Function cnBool(ByVal creditnote As DataSet) As Boolean
        If creditnote.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' check if DN exported
    ''' </summary>
    ''' <param name="creditnote"></param>
    ''' <returns></returns>
    Private Function dnBool(ByVal debitnote As DataSet) As Boolean
        If debitnote.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function anBool(ByVal adjustment As DataSet) As Boolean
        If adjustment.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' check if exported
    ''' </summary>
    ''' <param name="exported"></param>
    ''' <returns></returns>
    Private Function exportedBool(ByVal exported As DataSet) As Boolean
        If exported.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function equalDates(ByVal exported As DataSet) As Boolean
        If exported.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function chkIfAllExist(ByVal codeParam As String) As Boolean
        If CheckIfExist.IfExists(codeParam) = False And codeParam <> "" Then
            Return False
        Else
            Return True
        End If
    End Function

    'FUNCTION to get the vat amount
    Private Function calculateVat(ByVal amount As Double, ByVal vatCode As Integer) As Double
        Dim dtVatPerc As Double = generaldataset.loadVATPerc(vatCode)

        Return Math.Round((amount * dtVatPerc / 100), 2, MidpointRounding.AwayFromZero)
    End Function

    Public Sub MakeItExported(ByVal DeliveryMergedIDParam As String)

        Dim SQL5 = "UPDATE easyenquiry.easyenquiry.invoices SET exported='1' WHERE deliveryMerged_ID='" & DeliveryMergedIDParam & "';"
        Try
            myCommand.Connection = conn
            conn.Close()
            conn.Open()
            myCommand.CommandText = SQL5
            myCommand.ExecuteNonQuery()
            conn.Close()
        Catch ex As Exception

        End Try
    End Sub

    Public Sub cnExported(ByVal cnNumber As String)
        If cnNumber.Contains("CN/") Then
            Dim SQL5 = "UPDATE easyenquiry.easyenquiry.creditnotes SET exported='1' WHERE creditNoteNumber='" & cnNumber & "';"
            Try
                myCommand.Connection = conn
                conn.Close()
                conn.Open()
                myCommand.CommandText = SQL5
                myCommand.ExecuteNonQuery()
                conn.Close()
            Catch ex As Exception

            End Try
        Else
            Dim SQL5 = "UPDATE easyenquiry.easyenquiry.creditnotes SET exported='1' WHERE creditNoteSupNumber='" & cnNumber & "';"
            Try
                myCommand.Connection = conn
                conn.Close()
                conn.Open()
                myCommand.CommandText = SQL5
                myCommand.ExecuteNonQuery()
                conn.Close()
            Catch ex As Exception

            End Try
        End If
    End Sub

    Public Sub dnExported(ByVal dnNumber As String)

        Dim SQL5 = "UPDATE easyenquiry.easyenquiry.debitnotes SET exported='1' WHERE debitNoteNumber='" & dnNumber & "';"
        Try
            myCommand.Connection = conn
            conn.Close()
            conn.Open()
            myCommand.CommandText = SQL5
            myCommand.ExecuteNonQuery()
            conn.Close()
        Catch ex As Exception

        End Try
    End Sub
    Public Sub anExported(ByVal anID As String)

        Dim SQL5 = "UPDATE easyenquiry.easyenquiry.adjustments SET exported='1' WHERE adjustmentID='" & anID & "';"

        Try
            myCommand.Connection = conn
            conn.Close()
            conn.Open()
            myCommand.CommandText = SQL5
            myCommand.ExecuteNonQuery()
            conn.Close()
        Catch ex As Exception

        End Try
    End Sub

    Function datesMatch(ByVal datasetDates As DataSet) As Boolean
        Dim bool = True

        If datasetDates.Tables(0).Rows.Count = 0 Then
            Return bool
        Else
            For i = 0 To datasetDates.Tables(0).Rows.Count - 1
                If datasetDates.Tables(0).Rows(i).Item("invoiceDate") = datasetDates.Tables(0).Rows(i).Item("creditNoteDate") Then
                    bool = True
                Else
                    bool = False
                    Return bool
                End If
            Next
            Return bool
        End If
    End Function

    Function supplierInvoices(ByVal supinv As DataSet) As Boolean
        Dim bool = True
        Try
            If supinv.Tables(0).Rows.Count = 0 Then
                Return bool
            Else

                For i = 0 To supinv.Tables(0).Rows.Count - 1
                    If IsDBNull(supinv.Tables(0).Rows(i).Item("deliveryMergedID")) Then

                    Else
                        If supinv.Tables(0).Rows(i).Item("supplierInvoiceNumber") = "" Then
                            bool = False
                            Return bool
                        Else
                            bool = True
                        End If
                    End If
                Next

                Return bool
            End If
        Catch ex As Exception
            MessageBox.Show("One of the deliveries is missing a Supplier's Invoice Number.", "Important Message", MessageBoxButtons.OK)
        End Try
    End Function
End Class