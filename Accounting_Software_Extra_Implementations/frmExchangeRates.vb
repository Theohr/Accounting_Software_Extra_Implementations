Imports System.Data.SqlClient
Imports System.Data
Imports System.Net
Imports System.IO
Imports System.Xml
Imports System.Text
Imports mshtml
Imports System.Web

Public Class frmExchangeRates
    Dim myCommand As New SqlCommand
    Dim myAdapter As New SqlDataAdapter
    Dim SQL As String
    Dim myReader As SqlDataReader
    Dim exchangeRateID As String

    Dim pageLoad As Boolean = False
    Dim baseRateCurrencyCode As String = "USD"

    Dim XmlReader As XmlTextReader
    Dim rateAED As String = ""

    Dim webtext As String = ""

    Private Sub frmExchangeRates_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        setCurrencies()
        pageLoad = True
        LinkLabel1.Links(0).LinkData = "http://www.ecb.int/stats/exchange/eurofxref/html/index.en.html"
        'DateDailyExchangeRateDate.Checked = True




    End Sub

    Public Sub getAED()
        'XmlReader = New XmlTextReader("https://www.fx-exchange.com/aed/usd.xml")

        'While XmlReader.Read
        '    Console.WriteLine(XmlReader)
        'End While

        'XmlReader.Close()
        ''rateAED = ExpJSON.readCurJSON("USDAED")
        rateAED = ParseRssFile()
        'rateAED = "3.673"




    End Sub


    'Function parseMyHtml(ByVal htmlToParse$) As String
    '    Dim htmlDocument As IHTMLDocument2 = New HTMLDocumentClass()
    '    htmlDocument.write(htmlToParse)
    '    htmlDocument.close()

    '    Dim allElements As IHTMLElementCollection = htmlDocument.body.all

    '    Dim allInputs As IHTMLElementCollection = allElements.tags("a")
    '    Dim element As IHTMLElement
    '    For Each element In allInputs
    '        element.title = element.innerText
    '    Next

    '    Return htmlDocument.body.innerHTML
    'End Function
    Private Function ParseRssFile() As String
        Dim rssXmlDoc As XmlDocument = New XmlDocument()
        Dim tmp1StrStart As Integer
        Dim tmp1StrEnd As Integer
        Dim tmp1Str As String
        Dim data As String
        Dim rate As String

        Dim command = "curl https://www.fx-exchange.com/aed/usd.xml"

        Dim myprocess As New Process
        Dim StartInfo As New System.Diagnostics.ProcessStartInfo
        StartInfo.FileName = "cmd" 'starts cmd window
        StartInfo.RedirectStandardInput = True
        StartInfo.RedirectStandardOutput = True
        StartInfo.UseShellExecute = False 'required to redirect
        StartInfo.CreateNoWindow = True '<---- creates no window, obviously
        myprocess.StartInfo = StartInfo
        myprocess.Start()
        Dim SR As System.IO.StreamReader = myprocess.StandardOutput
        Dim SW As System.IO.StreamWriter = myprocess.StandardInput
        SW.WriteLine(command) 'the command you wish to run.....
        SW.WriteLine("exit") 'exits command prompt window
        Dim output = SR.ReadToEnd 'returns results of the command window
        SW.Close()
        SR.Close()

        ' Parse the Items in the RSS file
        Dim rssNodes As XmlNodeList = rssXmlDoc.SelectNodes("rss/channel/item")

        Dim rssContent As StringBuilder = New StringBuilder()

        ' Iterate through the items in the RSS file
        Dim rssNode As XmlNode
        For Each rssNode In rssNodes
            Dim rssSubNode As XmlNode = rssNode.SelectSingleNode("title")
            Dim title As String = IIf(rssSubNode.InnerText <> Nothing, rssSubNode.InnerText, "")

            rssSubNode = rssNode.SelectSingleNode("link")
            Dim link As String
            link = IIf(rssSubNode.InnerText <> Nothing, rssSubNode.InnerText, "")

            rssSubNode = rssNode.SelectSingleNode("description")
            Dim description As String = IIf(rssSubNode.InnerText <> Nothing, rssSubNode.InnerText, "")

            rssContent.Append("<a href='" + link + "'>" + title + "</a><br>" + description)
            data = description
        Next

        ' Return the string that contain the RSS items
        tmp1StrStart = output.IndexOf("1.00 USD =  ")
        If (tmp1StrStart > 0) Then
            tmp1StrEnd = output.IndexOf(" AED<br/>")
            tmp1Str = output.Substring(tmp1StrStart + 12, (tmp1StrEnd) - (tmp1StrStart + 12))
            rate = Trim(Replace(tmp1Str, vbTab, ""))
        Else
            rate = ""
        End If

        System.Console.WriteLine(rate)
        Return rate
    End Function



    Public Sub setCurrencies()
        Dim myData As New DataTable
        SQL = "SELECT * FROM easyenquiry.easyenquiry.currencies WHERE domesticCurrency=0"

        Try

            conn.Close()
            conn.Open()
            myCommand.Connection = conn
            myCommand.CommandText = SQL

            myAdapter.SelectCommand = myCommand
            myAdapter.Fill(myData)
            DataGridDailyExchangeRates.DataSource = myData

        Catch myerror As SqlException
            MsgBox("There was an error reading from the database: " & myerror.Message)
        End Try

    End Sub



    Public Sub setExchangeRates(ByVal tmpDailyExchangeRateDate As String)
        setCurrencies()
        Dim dataGridNumOfRows = DataGridDailyExchangeRates.RowCount - 1
        Dim tmpCurrencyCode As String
        Dim exchangeRateValue As String
        Dim rowCounter As Integer

        Dim currencyCode1 As String
        Dim currencyCode2 As String
        Dim currencyCode3 As String
        Dim currencyCode4 As String


        Try

            SQL = "SELECT currencies.currencyCode, exchangerates.exchangeRateValue FROM easyenquiry.easyenquiry.exchangerates LEFT JOIN easyenquiry.easyenquiry.currencies ON exchangerates.exchangeRateCurCode=currencies.currencyCode WHERE exchangerates.exchangeRateDate='" & tmpDailyExchangeRateDate & "'"
            myCommand.Connection = conn
            conn.Close()
            conn.Open()
            myCommand.CommandText = SQL
            myReader = myCommand.ExecuteReader
            If (myReader.HasRows = True) Then
                While myReader.Read


                    tmpCurrencyCode = myReader.GetValue(myReader.GetOrdinal("currencyCode"))
                    exchangeRateValue = myReader.GetValue(myReader.GetOrdinal("exchangeRateValue"))



                    For rowCounter = 0 To dataGridNumOfRows - 1

                        'currencyCode = DataGridDailyExchangeRates.Rows(rowCounter).Cells(0).Value
                        currencyCode1 = DataGridDailyExchangeRates.CurrentRow.Cells(0).Value
                        currencyCode2 = DataGridDailyExchangeRates.CurrentRow.Cells(1).Value
                        'currencyCode3 = DataGridDailyExchangeRates.CurrentRow.Cells(2).Value
                        currencyCode3 = DataGridDailyExchangeRates.Rows(rowCounter).Cells(2).Value
                        currencyCode4 = DataGridDailyExchangeRates.Rows(rowCounter).Cells(3).Value

                        System.Console.WriteLine("code1=" & currencyCode1)
                        System.Console.WriteLine("code2=" & currencyCode2)
                        System.Console.WriteLine("code3=" & currencyCode3)
                        System.Console.WriteLine("code4=" & currencyCode4)

                        If (currencyCode3 = tmpCurrencyCode) Then
                            DataGridDailyExchangeRates.Rows(rowCounter).Cells(0).Value = exchangeRateValue

                        End If







                    Next rowCounter




                End While
                myReader.Close()
            End If
        Catch ex As System.Exception

        End Try






    End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        Dim strDateDailyExchangeRateDate As String = ""
        Dim counter As Integer = 0
        If (DateDailyExchangeRateDate.Checked = True) Then
            DateDailyExchangeRateDate.CustomFormat = "yyyy-MM-dd"

            strDateDailyExchangeRateDate = DateDailyExchangeRateDate.Text


            Dim dailyExchangeRateDate As String
            Dim myDateTime As Date = Date.Now
            dailyExchangeRateDate = Format(myDateTime, "yyyy-MM-dd")

            Try
                SQL = "SELECT exchangeRateID FROM easyenquiry.easyenquiry.exchangerates WHERE exchangeRateDate='" & strDateDailyExchangeRateDate & "'"
                myCommand.Connection = conn
                conn.Close()
                conn.Open()
                myCommand.CommandText = SQL
                myReader = myCommand.ExecuteReader
                If (myReader.HasRows = True) Then
                    While myReader.Read
                        counter = counter + 1
                    End While
                    myReader.Close()
                End If
            Catch ex As System.Exception

            End Try

            'If ((strDateDailyExchangeRateDate < dailyExchangeRateDate) And (counter > 0)) Then
            If ((strDateDailyExchangeRateDate < dailyExchangeRateDate)) Then
                MsgBox("You cannot update the exchange rates since the date has been already Passed", MsgBoxStyle.Critical, "Date Cannot be modified")
            Else
                Dim rowCounter As Integer
                Dim dataGridNumOfRows = DataGridDailyExchangeRates.RowCount - 1

                Dim exchangeRateValue As String = ""
                'Dim dailyExchangeRateCurID As String = ""
                Dim currencyCode As String

                For rowCounter = 0 To dataGridNumOfRows - 1
                    exchangeRateValue = DataGridDailyExchangeRates.Rows(rowCounter).Cells(0).Value

                    If DataGridDailyExchangeRates.Rows(rowCounter).Cells(2).Value = "AED" Then
                        Try
                            If rateAED <> "" Then
                                DataGridDailyExchangeRates.Rows(rowCounter).Cells(0).Value = rateAED
                            End If
                        Catch ex As Exception
                            DataGridDailyExchangeRates.Rows(rowCounter).Cells(0).Value = rateAED
                        End Try
                    End If

                    If (exchangeRateValue = "") Then
                        MsgBox("You must supply a value for all currencies", MsgBoxStyle.Exclamation, "Null value found")
                        Exit Sub
                    End If

                Next

                For rowCounter = 0 To dataGridNumOfRows - 1

                    exchangeRateValue = DataGridDailyExchangeRates.Rows(rowCounter).Cells(0).Value
                    currencyCode = DataGridDailyExchangeRates.Rows(rowCounter).Cells(2).Value


                    System.Console.WriteLine("currencyCode=" & currencyCode)
                    System.Console.WriteLine("exchangeRateValue=" & exchangeRateValue)



                    Try
                        SQL = "SELECT * FROM easyenquiry.easyenquiry.exchangerates WHERE exchangeRateDate='" & strDateDailyExchangeRateDate & "' AND exchangeRateCurCode='" & currencyCode & "'"
                        myCommand.Connection = conn
                        conn.Close()
                        conn.Open()
                        myCommand.CommandText = SQL
                        myReader = myCommand.ExecuteReader
                        If (myReader.HasRows = True) Then
                            While myReader.Read
                                exchangeRateID = myReader.GetValue(myReader.GetOrdinal("exchangeRateID"))

                            End While
                            myReader.Close()
                        End If
                    Catch ex As System.Exception

                    End Try

                    If (exchangeRateID = "") Then
                        SQL = "INSERT INTO easyenquiry.easyenquiry. exchangerates(exchangeRateDate,exchangeRateCurCode,exchangeRateValue) VALUES(@exchangeRateDate,@exchangeRateCurCode,@exchangeRateValue)"
                    Else
                        SQL = "UPDATE  easyenquiry.easyenquiry.exchangerates SET exchangeRateValue=@exchangeRateValue WHERE exchangeRateDate='" & dailyExchangeRateDate & "' AND exchangeRateCurCode='" & currencyCode & "'"
                    End If
















                    Try
                        myCommand.Parameters.AddWithValue("@exchangeRateDate", strDateDailyExchangeRateDate)
                        myCommand.Parameters.AddWithValue("@exchangeRateCurCode", currencyCode)
                        myCommand.Parameters.AddWithValue("@exchangeRateValue", exchangeRateValue)

                        conn.Close()
                        conn.Open()
                        myCommand.CommandText = SQL
                        myCommand.ExecuteNonQuery()
                        myCommand.Parameters.Clear()


                    Catch ex As System.Exception

                    End Try




                    exchangeRateID = ""

                Next rowCounter



                Try

                    SQL = "SELECT * FROM easyenquiry.easyenquiry.exchangerates WHERE exchangeRateDate='" & strDateDailyExchangeRateDate & "' AND exchangeRateCurCode='" & baseRateCurrencyCode & "'"
                    myCommand.Connection = conn
                    conn.Close()
                    conn.Open()
                    myCommand.CommandText = SQL
                    myReader = myCommand.ExecuteReader
                    If (myReader.HasRows = True) Then
                        While myReader.Read
                            exchangeRateID = myReader.GetValue(myReader.GetOrdinal("exchangeRateID"))

                        End While
                        myReader.Close()
                    End If
                Catch ex As System.Exception

                End Try




                If (exchangeRateID = "") Then
                    Try
                        SQL = "INSERT INTO easyenquiry.easyenquiry. exchangerates(exchangeRateDate,exchangeRateCurCode,exchangeRateValue) VALUES(@exchangeRateDate,@exchangeRateCurCode,@exchangeRateValue)"
                        myCommand.Parameters.AddWithValue("@exchangeRateDate", strDateDailyExchangeRateDate)
                        myCommand.Parameters.AddWithValue("@exchangeRateCurCode", baseRateCurrencyCode)
                        myCommand.Parameters.AddWithValue("@exchangeRateValue", "1.00")

                        conn.Close()
                        conn.Open()
                        myCommand.CommandText = SQL
                        myCommand.ExecuteNonQuery()
                        myCommand.Parameters.Clear()


                    Catch ex As System.Exception

                    End Try

                End If



                MsgBox("Exchange Rates has been succesfully be updated", MsgBoxStyle.Information, "Exchange Rates")


            End If

        End If
        'myCommand.Parameters.AddWithValue("?dailyExchangeRateDate", strDateDailyExchangeRateDate)
        DateDailyExchangeRateDate.CustomFormat = "dd-MM-yyyy"
        Me.Close()


    End Sub


    Private Sub DateDailyExchangeRateDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateDailyExchangeRateDate.ValueChanged 'deftero
        Dim strDateDailyExchangeRateDate As String = ""
        DateDailyExchangeRateDate.CustomFormat = "yyyy-MM-dd"
        strDateDailyExchangeRateDate = DateDailyExchangeRateDate.Text
        DateDailyExchangeRateDate.CustomFormat = "dd-MM-yyyy"
        setExchangeRates(strDateDailyExchangeRateDate)
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub DataGridDailyExchangeRates_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridDailyExchangeRates.CellContentClick

    End Sub

    Private Sub DataGridDailyExchangeRates_CellLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridDailyExchangeRates.CellLeave

    End Sub

    Private Sub DataGridDailyExchangeRates_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridDailyExchangeRates.CellValueChanged
        Dim exchangeRateValueDbl As Double = 0.0
        Try
            'exchangeRateValueDec = CStr(CType(DataGridDailyExchangeRates.Rows(DataGridDailyExchangeRates.CurrentRow.Index).DataBoundItem, DataRowView).Row("exchangeRateValue"))
            If (CStr(DataGridDailyExchangeRates.CurrentRow.Cells(0).Value) <> "") Then
                exchangeRateValueDbl = CDbl(DataGridDailyExchangeRates.CurrentRow.Cells(0).Value)
                DataGridDailyExchangeRates.CurrentRow.Cells(0).Value = exchangeRateValueDbl
                System.Console.WriteLine("exchangeRateValue 0=" & exchangeRateValueDbl)
            End If


        Catch ex As System.Exception
            If (pageLoad = True) Then
                DataGridDailyExchangeRates.CurrentRow.Cells(0).Value = ""
                MsgBox("Please check the input rate, incompatible format number has been found", MsgBoxStyle.Exclamation, "Input Error")
            End If

        End Try
    End Sub

    Private Sub DataGridDailyExchangeRates_CurrentCellChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridDailyExchangeRates.CurrentCellChanged
        'Dim exchangeRateValueDec As Decimal = 0.0

    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim URL As String = "http://www.ecb.int/stats/exchange/eurofxref/html/index.en.html"
        Process.Start(URL)

    End Sub

    Private Sub btnGetRates_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetRates.Click
        'Dim page As String = GetPage("http://www.ecb.int/stats/exchange/eurofxref/html/index.en.html")
        'Dim body As String = ExtractBody(page)
        'System.Console.WriteLine(body)


        'Dim USD_Rate As String
        'Dim USD_Index As String
        'USD_Index = body.IndexOf("id=""USD""")



        'Dim exchangeRateValue As String = ""
        Dim currencyCode As String = ""


        Dim rowCounter As Integer = 0
        Dim dataGridNumOfRows = DataGridDailyExchangeRates.RowCount - 1

        Dim baseRateCurrencyRate As Double = 0.0


        Dim my_XmlReader As XmlReader


        getAED()



        'Try
        '    ''''' START GET PRICES WHEN THERE IS A PROXY SERVER '''''''''''''''''''''
        '    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create("http://www.ecb.int/stats/eurofxref/eurofxref-daily.xml")
        '    req.Proxy = New System.Net.WebProxy("http://192.168.200.4:8080", True)
        '    req.Proxy.Credentials = System.Net.CredentialCache.DefaultCredentials
        '    Dim resp As System.Net.WebResponse = req.GetResponse
        '    Dim textReader As StreamReader = New StreamReader(resp.GetResponseStream)
        '    my_XmlReader = New XmlTextReader(textReader)
        '    ''''' END GET PRICES WHEN THERE IS A PROXY SERVER '''''''''''''''''''''

        'Catch ex As System.Exception
        '    ''''' START GET PRICES WHEN THERE IS NO ANY PROXY SERVER '''''''''''''''''''''
        '    my_XmlReader = XmlReader.Create("http://www.ecb.int/stats/eurofxref/eurofxref-daily.xml")
        '    ''''' END GET PRICES WHEN THERE IS NO ANY PROXY SERVER '''''''''''''''''''''

        'End Try


        my_XmlReader = Xml.XmlReader.Create("https://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml")


        While my_XmlReader.Read()
            If ((my_XmlReader.NodeType = XmlNodeType.Element) And (my_XmlReader.Name = "Cube")) Then
                If (my_XmlReader.HasAttributes) Then
                    If (my_XmlReader.GetAttribute("rate") <> "") Then
                        'System.Console.WriteLine(my_XmlReader.GetAttribute("currency") + ": " + my_XmlReader.GetAttribute("rate"))
                        For rowCounter = 0 To dataGridNumOfRows - 1
                            'exchangeRateValue = DataGridDailyExchangeRates.Rows(rowCounter).Cells(0).Value
                            If (my_XmlReader.GetAttribute("currency") = "USD") Then
                                baseRateCurrencyRate = my_XmlReader.GetAttribute("rate")
                            End If

                            currencyCode = DataGridDailyExchangeRates.Rows(rowCounter).Cells(2).Value
                            If (currencyCode = my_XmlReader.GetAttribute("currency")) Then
                                DataGridDailyExchangeRates.Rows(rowCounter).Cells(0).Value = Math.Round((my_XmlReader.GetAttribute("rate") / baseRateCurrencyRate), 5)
                            ElseIf (currencyCode = "EUR") Then
                                DataGridDailyExchangeRates.Rows(rowCounter).Cells(0).Value = Math.Round((1 / baseRateCurrencyRate), 5)
                            ElseIf (currencyCode = "AED") Then
                                DataGridDailyExchangeRates.Rows(rowCounter).Cells(0).Value = rateAED
                            End If
                        Next

                    End If

                End If
            End If


        End While










    End Sub

    Private Sub btnCancel_Click_1(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub


End Class