Imports System.Data.SqlClient

Public Class frmItemComments
    Dim myCommand As New SqlCommand
    Dim myAdapter As New SqlDataAdapter
    Dim SQL As String
    Dim myReader As SqlDataReader

    Dim enquiryID As String = ""
    Dim supQuotationID As String = ""
    Dim itemListID As String = ""
    Dim supQuotItemListID As String = ""
    Dim supEvalQuotID As String = ""

    Dim generaldataset As New datasetsAndTables()

    Public Sub setEnquiryID(ByVal enquiryIDParam As String)
        enquiryID = enquiryIDParam
    End Sub

    Public Sub setSupQuotationID(ByVal supQuotationIDParam As String)
        supQuotationID = supQuotationIDParam
    End Sub

    Public Sub setItemListID(ByVal itemListIDParam As String)
        itemListID = itemListIDParam
    End Sub

    Public Sub setSupQuotItemListID(ByVal supQuotItemListIDParam As String)
        supQuotItemListID = supQuotItemListIDParam
    End Sub

    Public Sub setSupEvalQuotID(ByVal supEvalQuotIDParam As String)
        supEvalQuotID = supEvalQuotIDParam
    End Sub

    Private Sub frmItemComments_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadItemComments()
    End Sub

    Public Sub loadItemComments()
        Dim comments = ""

        If itemListID <> "" Then
            SQL = "SELECT itemComments FROM easyenquiry.easyenquiry.itemslist WHERE itemListID='" & itemListID & "'"
        ElseIf supQuotItemListID <> "" Then
            SQL = "SELECT supQuotItemListComments FROM easyenquiry.easyenquiry.supquotitemslist WHERE supQuotItemListID='" & supQuotItemListID & "'"
        End If

        Try
            myCommand.Connection = conn
            conn.Close()
            conn.Open()
            myCommand.CommandText = SQL
            myReader = myCommand.ExecuteReader
            If (myReader.HasRows = True) Then
                While myReader.Read
                    If itemListID <> "" Then
                        comments = myReader.GetValue(myReader.GetOrdinal("itemComments"))
                    ElseIf supQuotItemListID <> "" Then
                        comments = myReader.GetValue(myReader.GetOrdinal("supQuotItemListComments"))
                    End If
                End While
                myReader.Close()
            End If
        Catch ex As Exception

        End Try

        txtItemComments.Text = comments
    End Sub

    Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click
        If itemListID <> "" Then
            SQL = "UPDATE easyenquiry.easyenquiry.itemslist SET itemComments=@itemComments WHERE itemListID='" & itemListID & "'"
            myCommand.Parameters.AddWithValue("@itemComments", txtItemComments.Text.Trim())
        ElseIf supQuotItemListID <> "" Then
            SQL = "UPDATE easyenquiry.easyenquiry.supquotitemslist SET supQuotItemListComments=@supQuotItemListComments WHERE supQuotItemListID='" & supQuotItemListID & "'"
            myCommand.Parameters.AddWithValue("@supQuotItemListComments", txtItemComments.Text.Trim())
        End If

        Try
            conn.Close()
            conn.Open()
            myCommand.CommandText = SQL
            myCommand.ExecuteNonQuery()
            myCommand.Parameters.Clear()
            Try
                Dim frm As Form
                Dim egwEperasa = "Eperasa"
                For Each frm In My.Application.OpenForms
                    If (frm Is My.Forms.frmSupplierQuotation) Then
                        If itemListID <> "" Then
                            frmSupplierQuotation.changeCommentButtonColor(itemListID)
                        ElseIf supQuotItemListID <> "" Then
                            frmSupplierQuotation.changeCommentButtonColor(supQuotItemListID)
                        End If
                        Me.Close()
                    ElseIf (frm Is My.Forms.frmCustomerQuotations) Then
                        'frmCustomerQuotations.loadItemsList(supEvalQuotID, egwEperasa)
                        frmCustomerQuotations.changeCommentButtonColor(supQuotItemListID)
                        Me.Close()
                    ElseIf (frm Is My.Forms.frmDelivery) Then
                        frmDelivery.changeCommentButtonColor(supQuotItemListID)
                        Me.Close()
                    End If
                Next
            Catch ex As Exception

            End Try
            'MessageBox.Show("Comments inserted Succesfully!", "Important Message", MessageBoxButtons.OK)
        Catch myerror As SqlException
            MessageBox.Show("There was an error in commenting the specific item. Please contact your system administrator", "Important Message", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

End Class