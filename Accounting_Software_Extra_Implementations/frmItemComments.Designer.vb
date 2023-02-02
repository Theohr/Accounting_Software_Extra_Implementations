<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmItemComments
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lblComments = New System.Windows.Forms.Label()
        Me.txtItemComments = New System.Windows.Forms.RichTextBox()
        Me.btnOk = New ComponentFactory.Krypton.Toolkit.KryptonButton()
        Me.btnCancel = New ComponentFactory.Krypton.Toolkit.KryptonButton()
        Me.SuspendLayout()
        '
        'lblComments
        '
        Me.lblComments.AutoSize = True
        Me.lblComments.BackColor = System.Drawing.SystemColors.Control
        Me.lblComments.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblComments.Location = New System.Drawing.Point(214, 9)
        Me.lblComments.Name = "lblComments"
        Me.lblComments.Size = New System.Drawing.Size(94, 20)
        Me.lblComments.TabIndex = 0
        Me.lblComments.Text = "Comments"
        '
        'txtItemComments
        '
        Me.txtItemComments.Location = New System.Drawing.Point(12, 40)
        Me.txtItemComments.Name = "txtItemComments"
        Me.txtItemComments.Size = New System.Drawing.Size(511, 228)
        Me.txtItemComments.TabIndex = 1
        Me.txtItemComments.Text = ""
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(12, 274)
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
        Me.btnOk.TabIndex = 217
        Me.btnOk.Values.Image = Global.easyEnquiry.My.Resources.Resources.ok
        Me.btnOk.Values.Text = "Save"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(428, 274)
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
        Me.btnCancel.TabIndex = 216
        Me.btnCancel.Values.Image = Global.easyEnquiry.My.Resources.Resources.close
        Me.btnCancel.Values.Text = "Cancel"
        '
        'frmItemComments
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(535, 313)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.txtItemComments)
        Me.Controls.Add(Me.lblComments)
        Me.Name = "frmItemComments"
        Me.Text = "frmItemComments"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblComments As Label
    Friend WithEvents txtItemComments As RichTextBox
    Private WithEvents btnOk As ComponentFactory.Krypton.Toolkit.KryptonButton
    Private WithEvents btnCancel As ComponentFactory.Krypton.Toolkit.KryptonButton
End Class
