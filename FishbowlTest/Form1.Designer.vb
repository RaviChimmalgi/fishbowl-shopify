<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.GetShopifyOrders = New System.Windows.Forms.Button()
        Me.PutOrdersInFishBowl = New System.Windows.Forms.Button()
        Me.GetFishBowlOrders = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.StatusStrip1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(35, 53)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.TextBox1.Size = New System.Drawing.Size(348, 350)
        Me.TextBox1.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(35, 24)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 406)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(960, 22)
        Me.StatusStrip1.TabIndex = 2
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(0, 17)
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(402, 53)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.Size = New System.Drawing.Size(546, 350)
        Me.DataGridView1.TabIndex = 3
        '
        'GetShopifyOrders
        '
        Me.GetShopifyOrders.Location = New System.Drawing.Point(341, 24)
        Me.GetShopifyOrders.Name = "GetShopifyOrders"
        Me.GetShopifyOrders.Size = New System.Drawing.Size(117, 23)
        Me.GetShopifyOrders.TabIndex = 4
        Me.GetShopifyOrders.Text = "GetShopifyOrders"
        Me.GetShopifyOrders.UseVisualStyleBackColor = True
        '
        'PutOrdersInFishBowl
        '
        Me.PutOrdersInFishBowl.Location = New System.Drawing.Point(501, 24)
        Me.PutOrdersInFishBowl.Name = "PutOrdersInFishBowl"
        Me.PutOrdersInFishBowl.Size = New System.Drawing.Size(147, 23)
        Me.PutOrdersInFishBowl.TabIndex = 5
        Me.PutOrdersInFishBowl.Text = "PutOrdersInFishBowl"
        Me.PutOrdersInFishBowl.UseVisualStyleBackColor = True
        '
        'GetFishBowlOrders
        '
        Me.GetFishBowlOrders.Location = New System.Drawing.Point(183, 24)
        Me.GetFishBowlOrders.Name = "GetFishBowlOrders"
        Me.GetFishBowlOrders.Size = New System.Drawing.Size(109, 23)
        Me.GetFishBowlOrders.TabIndex = 6
        Me.GetFishBowlOrders.Text = "GetFishBowlOrders"
        Me.GetFishBowlOrders.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(699, 24)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(141, 23)
        Me.Button2.TabIndex = 7
        Me.Button2.Text = "UpdateShopifyQuantity"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(960, 428)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.GetFishBowlOrders)
        Me.Controls.Add(Me.PutOrdersInFishBowl)
        Me.Controls.Add(Me.GetShopifyOrders)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents GetShopifyOrders As System.Windows.Forms.Button
    Friend WithEvents PutOrdersInFishBowl As System.Windows.Forms.Button
    Friend WithEvents GetFishBowlOrders As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button

End Class
