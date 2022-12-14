Public Class ucComboboxAutoCompleteCell
    Inherits DataGridViewComboBoxCell

    Public Sub New()
        MyBase.New()
        DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox
        FlatStyle = FlatStyle.Flat
    End Sub
    Public Overrides ReadOnly Property EditType As Type
        Get
            Return GetType(ucComboboxAutoComplete)
        End Get
    End Property
    Public Overrides Sub InitializeEditingControl(ByVal rowIndex As Integer,
        ByVal initialFormattedValue As Object,
        ByVal dataGridViewCellStyle As DataGridViewCellStyle)
        ' Call base...
        MyBase.InitializeEditingControl(rowIndex, initialFormattedValue, dataGridViewCellStyle)

        Dim ctl As ucComboboxAutoComplete = CType(DataGridView.EditingControl, ucComboboxAutoComplete)
        ctl.DropDownStyle = ComboBoxStyle.DropDown
        ctl.AutoCompleteMode = AutoCompleteMode.None

        ' Make sure you have an instance...
        If ctl IsNot Nothing Then
            Dim ucColumn As ucComboboxAutoCompleteColumn
            ucColumn = TryCast(OwningColumn, ucComboboxAutoCompleteColumn)

            If ucColumn IsNot Nothing Then
                ctl.BeginUpdate()
                ctl.AllowForItemsOnly = True
                ctl.DropDownStyle = ComboBoxStyle.DropDown
                ctl.synonymMode_SetMode()
                For Each val As String In ucColumn.Items
                    If mActualText.Contains(val) Then
                        Dim idx = mActualText.IndexOf(val)
                        ctl.synonymMode_AddItem(mActualText(idx), mID(idx), mSynonymList(idx))
                    Else
                        ctl.synonymMode_AddItem(val, val)
                    End If
                Next
                ctl.Text = initialFormattedValue
                ctl.EndUpdate()
                ctl.showDropDown()
            End If
        End If

    End Sub

    Private mActualText As New List(Of String)
    Private mID As New List(Of String)
    Private mSynonymList As New List(Of String)
    Private mbSetted As Boolean = True
    Public Sub SetData(ByRef sActualText As String, sID As String, sSynonymList As String)
        Dim ucColumn As ucComboboxAutoCompleteColumn
        ucColumn = TryCast(OwningColumn, ucComboboxAutoCompleteColumn)

        If ucColumn IsNot Nothing Then
            If Not ucColumn.Items.Contains(sActualText) Then
                ucColumn.Items.Add(sActualText)
            End If
        End If
        mActualText.Add(sActualText)
        mID.Add(sID)
        mSynonymList.Add(sSynonymList)
        Me.Value = sActualText
    End Sub

End Class
