Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        UcComboboxAutoComplete1.synonymMode = True
        UcComboboxAutoComplete1.synonymMode_AddItem("Hello", "1", "Bob")
        UcComboboxAutoComplete1.synonymMode_AddItem("Chicken", "2", "Jack")

        UcComboboxAutoComplete2.synonymMode = True
        UcComboboxAutoComplete2.synonymMode_AddItem("<aaaHello1", "1", "a")
        UcComboboxAutoComplete2.synonymMode_AddItem("Hello2", "2", "b")
        UcComboboxAutoComplete2.synonymMode_AddItem("Hello3", "3", "c")
        UcComboboxAutoComplete2.synonymMode_AddItem("Hello4", "4", "d")

        DataGridView1.Rows.Add()
        DataGridView1.Rows.Add()
        DataGridView1.Rows.Add()
        DataGridView1.Rows.Add()
        DataGridView1.Rows.Add()

        DataGridView1.Rows(0).Cells(0) = New ucComboboxAutoCompleteCell()
        DataGridView1.Rows(1).Cells(0) = New ucComboboxAutoCompleteCell()
        DataGridView1.Rows(2).Cells(0) = New ucComboboxAutoCompleteCell()

        Dim cb1 As DataGridViewComboBoxCell = New DataGridViewComboBoxCell()
        cb1.Items.Add("13")
        cb1.Items.Add("abx")
        cb1.Items.Add("sdf")
        DataGridView1.Rows(0).Cells(1) = cb1
        DataGridView1.Rows(1).Cells(1) = New ucComboboxAutoCompleteCell()
        DataGridView1.Rows(2).Cells(1) = New DataGridViewComboBoxCell()
    End Sub

    Private Sub UcComboboxAutoComplete1_NewItemAdded(value As String, tag As String, ByRef sender As ucComboboxAutoComplete) Handles UcComboboxAutoComplete1.NewItemAdded
        MessageBox.Show("Value:" + value + " Tag: " + tag, "New Item added")
    End Sub
End Class
