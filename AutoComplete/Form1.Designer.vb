<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
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
        Me.UcComboboxAutoComplete1 = New AutoComplete.ucComboboxAutoComplete()
        Me.UcComboboxAutoComplete2 = New AutoComplete.ucComboboxAutoComplete()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Column1 = New AutoComplete.ucComboboxAutoCompleteColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewComboBoxColumn()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'UcComboboxAutoComplete1
        '
        Me.UcComboboxAutoComplete1.AllowForItemsOnly = False
        Me.UcComboboxAutoComplete1.AllowUserToAdd = False
        Me.UcComboboxAutoComplete1.EditingControlDataGridView = Nothing
        Me.UcComboboxAutoComplete1.EditingControlFormattedValue = ""
        Me.UcComboboxAutoComplete1.EditingControlRowIndex = 0
        Me.UcComboboxAutoComplete1.EditingControlValueChanged = False
        Me.UcComboboxAutoComplete1.FormattingEnabled = True
        Me.UcComboboxAutoComplete1.Location = New System.Drawing.Point(209, 126)
        Me.UcComboboxAutoComplete1.Name = "UcComboboxAutoComplete1"
        Me.UcComboboxAutoComplete1.SelectAllAutomatically = True
        Me.UcComboboxAutoComplete1.Size = New System.Drawing.Size(300, 23)
        Me.UcComboboxAutoComplete1.Sorted = True
        Me.UcComboboxAutoComplete1.TabIndex = 0
        '
        'UcComboboxAutoComplete2
        '
        Me.UcComboboxAutoComplete2.AllowForItemsOnly = False
        Me.UcComboboxAutoComplete2.EditingControlDataGridView = Nothing
        Me.UcComboboxAutoComplete2.EditingControlFormattedValue = ""
        Me.UcComboboxAutoComplete2.EditingControlRowIndex = 0
        Me.UcComboboxAutoComplete2.EditingControlValueChanged = False
        Me.UcComboboxAutoComplete2.FormattingEnabled = True
        Me.UcComboboxAutoComplete2.Location = New System.Drawing.Point(209, 155)
        Me.UcComboboxAutoComplete2.Name = "UcComboboxAutoComplete2"
        Me.UcComboboxAutoComplete2.SelectAllAutomatically = True
        Me.UcComboboxAutoComplete2.Size = New System.Drawing.Size(300, 23)
        Me.UcComboboxAutoComplete2.Sorted = True
        Me.UcComboboxAutoComplete2.TabIndex = 1
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(87, 201)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column2})
        Me.DataGridView1.Location = New System.Drawing.Point(220, 210)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(505, 150)
        Me.DataGridView1.TabIndex = 3
        '
        'Column1
        '
        Me.Column1.AllowUserToAdd = False
        Me.Column1.AllowForItemsOnly = False
        Me.Column1.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox
        Me.Column1.DropDownWidth = 160
        Me.Column1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Column1.HeaderText = "ucCustom"
        Me.Column1.Items.AddRange(New Object() {"Hello1", "Hello2", "Hello3", "Hello4", "Hello5", "Hello6", "Hello7"})
        Me.Column1.Name = "Column1"
        Me.Column1.Width = 160
        '
        'Column2
        '
        Me.Column2.HeaderText = "Default"
        Me.Column2.Items.AddRange(New Object() {"default1", "default2", "default3", "default4", "default5", "default6", "default7"})
        Me.Column2.Name = "Column2"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.UcComboboxAutoComplete2)
        Me.Controls.Add(Me.UcComboboxAutoComplete1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents UcComboboxAutoComplete1 As ucComboboxAutoComplete
    Friend WithEvents UcComboboxAutoComplete2 As ucComboboxAutoComplete
    Friend WithEvents Button1 As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Column1 As ucComboboxAutoCompleteColumn
    Friend WithEvents Column2 As DataGridViewComboBoxColumn
End Class
