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

        For i As Integer = 1 To 7
            Dim row As String() = New String() {String.Format("Hello{0}", i), String.Format("default{0}", i)}
            DataGridView1.Rows.Add(row)
        Next

        Dim Autocompletecell1 As ucComboboxAutoCompleteCell = DataGridView1.Rows(0).Cells(0)
        Autocompletecell1.SetData("Value1", "id1", "synonym1")
        Autocompletecell1.SetData("Value2", "id2", "synonym2")
        Autocompletecell1.SetData("Value3", "id3", "synonym3")

    End Sub

    Private Sub UcComboboxAutoComplete1_NewItemAdded(value As String, tag As String, ByRef sender As ucComboboxAutoComplete) Handles UcComboboxAutoComplete1.NewItemAdded
        MessageBox.Show("Value:" + value + " Tag: " + tag, "New Item added")
    End Sub
End Class
