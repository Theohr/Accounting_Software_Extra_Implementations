<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmExchangeRates
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmExchangeRates))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.DateDailyExchangeRateDate = New System.Windows.Forms.DateTimePicker()
        Me.DataGridDailyExchangeRates = New System.Windows.Forms.DataGridView()
        Me.currencyID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.domesticCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.currencyCode = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.exchangeRateValue = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel()
        Me.btnGetRates = New System.Windows.Forms.Button()
        Me.btnOk = New ComponentFactory.Krypton.Toolkit.KryptonButton()
        Me.btnCancel = New ComponentFactory.Krypton.Toolkit.KryptonButton()
        CType(Me.DataGridDailyExchangeRates, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(70, 59)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 198
        Me.Label1.Text = "Date"
        '
        'DateDailyExchangeRateDate
        '
        Me.DateDailyExchangeRateDate.Checked = False
        Me.DateDailyExchangeRateDate.CustomFormat = "dd-MM-yyy"
        Me.DateDailyExchangeRateDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateDailyExchangeRateDate.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.DateDailyExchangeRateDate.Location = New System.Drawing.Point(103, 56)
        Me.DateDailyExchangeRateDate.Name = "DateDailyExchangeRateDate"
        Me.DateDailyExchangeRateDate.ShowCheckBox = True
        Me.DateDailyExchangeRateDate.Size = New System.Drawing.Size(139, 20)
        Me.DateDailyExchangeRateDate.TabIndex = 197
        '
        'DataGridDailyExchangeRates
        '
        Me.DataGridDailyExchangeRates.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridDailyExchangeRates.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.currencyID, Me.domesticCurrency, Me.currencyCode, Me.exchangeRateValue})
        Me.DataGridDailyExchangeRates.Location = New System.Drawing.Point(37, 82)
        Me.DataGridDailyExchangeRates.Name = "DataGridDailyExchangeRates"
        Me.DataGridDailyExchangeRates.Size = New System.Drawing.Size(244, 305)
        Me.DataGridDailyExchangeRates.TabIndex = 196
        '
        'currencyID
        '
        Me.currencyID.DataPropertyName = "currencyID"
        DataGridViewCellStyle1.NullValue = Nothing
        Me.currencyID.DefaultCellStyle = DataGridViewCellStyle1
        Me.currencyID.HeaderText = "Currency ID"
        Me.currencyID.Name = "currencyID"
        Me.currencyID.Visible = False
        '
        'domesticCurrency
        '
        Me.domesticCurrency.DataPropertyName = "domesticCurrency"
        Me.domesticCurrency.HeaderText = "Domestic Currency"
        Me.domesticCurrency.Name = "domesticCurrency"
        Me.domesticCurrency.Visible = False
        '
        'currencyCode
        '
        Me.currencyCode.DataPropertyName = "currencyCode"
        Me.currencyCode.HeaderText = "CURRENCY"
        Me.currencyCode.Name = "currencyCode"
        '
        'exchangeRateValue
        '
        Me.exchangeRateValue.DataPropertyName = "exchangeRateValue"
        Me.exchangeRateValue.HeaderText = "RATE"
        Me.exchangeRateValue.Name = "exchangeRateValue"
        '
        'LinkLabel1
        '
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(76, 31)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(86, 13)
        Me.LinkLabel1.TabIndex = 201
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Exchange Rates"
        '
        'btnGetRates
        '
        Me.btnGetRates.Location = New System.Drawing.Point(167, 27)
        Me.btnGetRates.Name = "btnGetRates"
        Me.btnGetRates.Size = New System.Drawing.Size(75, 23)
        Me.btnGetRates.TabIndex = 202
        Me.btnGetRates.Text = "Get Rates"
        Me.btnGetRates.UseVisualStyleBackColor = True
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(59, 417)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black
        Me.btnOk.Size = New System.Drawing.Size(95, 30)
        Me.btnOk.StateNormal.Content.ShortText.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOk.StatePressed.Back.Color1 = System.Drawing.Color.DimGray
        Me.btnOk.StatePressed.Back.Color2 = System.Drawing.Color.DimGray
        Me.btnOk.StatePressed.Border.Color1 = System.Drawing.Color.DimGray
        Me.btnOk.StatePressed.Border.Color2 = System.Drawing.Color.DarkGray
        Me.btnOk.StatePressed.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.btnOk.StatePressed.Content.ShortText.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOk.StateTracking.Back.Color1 = System.Drawing.Color.DimGray
        Me.btnOk.StateTracking.Back.Color2 = System.Drawing.Color.DimGray
        Me.btnOk.StateTracking.Border.Color1 = System.Drawing.Color.DimGray
        Me.btnOk.StateTracking.Border.Color2 = System.Drawing.Color.DarkGray
        Me.btnOk.StateTracking.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.btnOk.StateTracking.Content.ShortText.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOk.TabIndex = 215
        Me.btnOk.Values.Image = Global.easyEnquiry.My.Resources.Resources.ok
        Me.btnOk.Values.Text = "Ok"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(160, 417)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black
        Me.btnCancel.Size = New System.Drawing.Size(95, 30)
        Me.btnCancel.StateCommon.Content.ShortText.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.StateNormal.Content.ShortText.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.StatePressed.Back.Color1 = System.Drawing.Color.DimGray
        Me.btnCancel.StatePressed.Back.Color2 = System.Drawing.Color.DimGray
        Me.btnCancel.StatePressed.Border.Color1 = System.Drawing.Color.DimGray
        Me.btnCancel.StatePressed.Border.Color2 = System.Drawing.Color.DarkGray
        Me.btnCancel.StatePressed.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.btnCancel.StatePressed.Content.ShortText.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.StateTracking.Back.Color1 = System.Drawing.Color.DimGray
        Me.btnCancel.StateTracking.Back.Color2 = System.Drawing.Color.DimGray
        Me.btnCancel.StateTracking.Border.Color1 = System.Drawing.Color.DimGray
        Me.btnCancel.StateTracking.Border.Color2 = System.Drawing.Color.DarkGray
        Me.btnCancel.StateTracking.Border.DrawBorders = CType((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) _
            Or ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right), ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)
        Me.btnCancel.StateTracking.Content.ShortText.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.TabIndex = 214
        Me.btnCancel.Values.Image = Global.easyEnquiry.My.Resources.Resources.close
        Me.btnCancel.Values.Text = "Cancel"
        '
        'frmExchangeRates
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(317, 459)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnGetRates)
        Me.Controls.Add(Me.LinkLabel1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DateDailyExchangeRateDate)
        Me.Controls.Add(Me.DataGridDailyExchangeRates)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(333, 498)
        Me.MinimumSize = New System.Drawing.Size(333, 498)
        Me.Name = "frmExchangeRates"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmExchangeRates"
        CType(Me.DataGridDailyExchangeRates, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DateDailyExchangeRateDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents DataGridDailyExchangeRates As System.Windows.Forms.DataGridView
    Friend WithEvents currencyID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents domesticCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents currencyCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents exchangeRateValue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents btnGetRates As System.Windows.Forms.Button
    Friend WithEvents btnOk As ComponentFactory.Krypton.Toolkit.KryptonButton
    Friend WithEvents btnCancel As ComponentFactory.Krypton.Toolkit.KryptonButton
End Class
